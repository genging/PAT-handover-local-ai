from __future__ import annotations

import argparse
import pickle
from pathlib import Path

import faiss
import numpy as np
import torch
from PIL import Image
from transformers import AutoImageProcessor, AutoModel


MODEL_NAME = "facebook/dinov2-small"
APP_ROOT = Path(__file__).resolve().parent
REFERENCE_ROOT = APP_ROOT.parent / "0_Reference_Photos"
INDEX_ROOT = APP_ROOT / "dinov2_faiss"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
DEVICE = "cuda" if torch.cuda.is_available() else "cpu"
_MODEL = None
_PROCESSOR = None


def preprocess_for_dino(pil_image: Image.Image) -> Image.Image:
    """Keep the central 70% of the photo to reduce sky, frame and watermark bias."""
    image = pil_image.convert("RGB")
    width, height = image.size
    left, top = int(width * 0.15), int(height * 0.15)
    right, bottom = int(width * 0.85), int(height * 0.85)
    return image.crop((left, top, max(right, left + 1), max(bottom, top + 1)))


def _load_model():
    global _MODEL, _PROCESSOR
    if _MODEL is None:
        _PROCESSOR = AutoImageProcessor.from_pretrained(MODEL_NAME)
        _MODEL = AutoModel.from_pretrained(MODEL_NAME).to(DEVICE).eval()
    return _MODEL, _PROCESSOR


def _embed_paths(paths: list[Path], batch_size: int = 32, progress_callback=None) -> list[list[float]]:
    model, processor = _load_model()
    vectors: list[list[float]] = []
    with torch.inference_mode():
        for start in range(0, len(paths), batch_size):
            batch_paths = paths[start:start + batch_size]
            images = []
            valid_positions = []
            for position, path in enumerate(batch_paths):
                try:
                    with Image.open(path) as image:
                        images.append(preprocess_for_dino(image))
                    valid_positions.append(position)
                except Exception:
                    pass
            if images:
                inputs = processor(images=images, return_tensors="pt")
                inputs = {key: value.to(DEVICE) for key, value in inputs.items()}
                output = model(**inputs)
                encoded = output.last_hidden_state[:, 0, :]
                encoded = torch.nn.functional.normalize(encoded, p=2, dim=1).cpu().numpy().astype("float32")
                by_position = {position: encoded[index].tolist() for index, position in enumerate(valid_positions)}
            else:
                by_position = {}
            vectors.extend(by_position.get(position, []) for position in range(len(batch_paths)))
            current = min(start + len(batch_paths), len(paths))
            print(f"  DINOv2 {current}/{len(paths)}", flush=True)
            if progress_callback:
                progress_callback(current, len(paths))
    return vectors


def build_reference_embeddings(progress_callback=None) -> dict:
    INDEX_ROOT.mkdir(parents=True, exist_ok=True)
    summary = {"model": MODEL_NAME, "device": DEVICE, "phases": {}}
    for phase in ("Before", "After"):
        phase_root = REFERENCE_ROOT / phase
        paths = [path for path in phase_root.rglob("*") if path.is_file() and path.suffix.casefold() in IMAGE_EXTENSIONS] if phase_root.exists() else []
        vectors = _embed_paths(paths, progress_callback=(lambda current, total, phase=phase: progress_callback(phase, current, total)) if progress_callback else None)
        records = []
        valid_vectors = []
        for path, vector in zip(paths, vectors):
            if vector:
                valid_vectors.append(vector)
                records.append({"phase": phase, "slot": path.parent.name, "path": str(path)})
        dimension = len(valid_vectors[0]) if valid_vectors else 384
        index = faiss.IndexFlatL2(dimension)
        if valid_vectors:
            index.add(np.asarray(valid_vectors, dtype="float32"))
        index_path = INDEX_ROOT / f"{phase.lower()}_index.faiss"
        mapping_path = INDEX_ROOT / f"{phase.lower()}_mapping.pkl"
        faiss.write_index(index, str(index_path))
        mapping_path.write_bytes(pickle.dumps(records, protocol=pickle.HIGHEST_PROTOCOL))
        summary["phases"][phase] = {"images": len(records), "index": str(index_path), "mapping": str(mapping_path), "dimension": dimension}
        print(f"{phase}: {len(records)} DINOv2 vectors -> {index_path}", flush=True)
    (INDEX_ROOT / "summary.json").write_text(__import__("json").dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def get_top_matches(image_path: str | Path, phase: str, top_k: int = 3) -> list[dict]:
    index_path = INDEX_ROOT / f"{phase.lower()}_index.faiss"
    mapping_path = INDEX_ROOT / f"{phase.lower()}_mapping.pkl"
    if not index_path.exists() or not mapping_path.exists():
        return []
    query = _embed_paths([Path(image_path)], batch_size=1)[0]
    if not query:
        return []
    index = faiss.read_index(str(index_path))
    records = pickle.loads(mapping_path.read_bytes())
    if index.ntotal == 0:
        return []
    # Retrieve ten reference images, then vote by slot instead of trusting one image.
    distances, indices = index.search(np.asarray([query], dtype="float32"), min(index.ntotal, max(10, top_k * 4)))
    votes = {}
    for distance, row in zip(distances[0], indices[0]):
        if row < 0 or row >= len(records):
            continue
        record = records[int(row)]
        slot = record["slot"]
        similarity = float(1.0 / (1.0 + max(float(distance), 0.0)))
        bucket = votes.setdefault(slot, {"weight": 0.0, "best_distance": float(distance), "best_similarity": similarity, "path": record["path"], "images": 0})
        bucket["weight"] += similarity
        bucket["images"] += 1
        if distance < bucket["best_distance"]:
            bucket["best_distance"] = float(distance)
            bucket["best_similarity"] = similarity
            bucket["path"] = record["path"]
    ranked = sorted(votes.items(), key=lambda item: (item[1]["weight"], item[1]["best_similarity"]), reverse=True)
    total_weight = sum(item[1]["weight"] for item in ranked) or 1.0
    return [{"slot": slot, "distance": data["best_distance"], "similarity": float(data["weight"] / total_weight), "best_similarity": data["best_similarity"], "votes": data["images"], "path": data["path"]} for slot, data in ranked[:top_k]]


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--build", action="store_true")
    parser.add_argument("--image")
    parser.add_argument("--phase", default="After")
    args = parser.parse_args()
    if args.build:
        build_reference_embeddings()
    if args.image:
        print(get_top_matches(args.image, args.phase))
