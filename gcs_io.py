from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Optional

from google.cloud import storage


def is_gcs_uri(value: str) -> bool:
    return value.startswith("gs://")


@dataclass(frozen=True)
class GcsPath:
    bucket: str
    prefix: str


def parse_gcs_uri(uri: str) -> GcsPath:
    if not is_gcs_uri(uri):
        raise ValueError(f"Not a GCS URI: {uri}")
    rest = uri[5:]
    bucket, _, prefix = rest.partition("/")
    return GcsPath(bucket=bucket, prefix=prefix.strip("/"))


def format_gcs_uri(bucket: str, prefix: str) -> str:
    if prefix:
        return f"gs://{bucket}/{prefix.strip('/')}"
    return f"gs://{bucket}"


class GcsIO:
    def __init__(self, client: Optional[storage.Client] = None) -> None:
        self.client = client or storage.Client()

    def list(self, bucket: str, prefix: str) -> Iterable[str]:
        b = self.client.bucket(bucket)
        for blob in b.list_blobs(prefix=prefix):
            yield blob.name

    def list_with_sizes(self, bucket: str, prefix: str) -> Iterable[tuple[str, int]]:
        b = self.client.bucket(bucket)
        for blob in b.list_blobs(prefix=prefix):
            yield blob.name, int(blob.size or 0)

    def exists(self, bucket: str, object_name: str) -> bool:
        return self.client.bucket(bucket).blob(object_name).exists()

    def upload_file(self, local_path: str, bucket: str, object_name: str) -> None:
        self.client.bucket(bucket).blob(object_name).upload_from_filename(local_path)

    def download_to_file(self, bucket: str, object_name: str, local_path: str) -> None:
        self.client.bucket(bucket).blob(object_name).download_to_filename(local_path)

    def write_text(self, bucket: str, object_name: str, text: str) -> None:
        self.client.bucket(bucket).blob(object_name).upload_from_string(
            text, content_type="text/plain"
        )

    def delete(self, bucket: str, object_name: str) -> None:
        self.client.bucket(bucket).blob(object_name).delete()

    def delete_prefix(self, bucket: str, prefix: str) -> None:
        b = self.client.bucket(bucket)
        for blob in b.list_blobs(prefix=prefix):
            blob.delete()
