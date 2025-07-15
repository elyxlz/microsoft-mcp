import httpx
import time
from typing import Any, Iterator
from .auth import get_token

BASE_URL = "https://graph.microsoft.com/v1.0"
# 15 x 320 KiB = 4,915,200 bytes
UPLOAD_CHUNK_SIZE = 15 * 320 * 1024

_client = httpx.Client(timeout=30.0, follow_redirects=True)


def _enhance_query_for_semantic_search(query: str) -> str:
    """Enhance query for better semantic understanding in Microsoft Graph Search"""
    # Remove common search operators that might interfere with semantic search
    query = query.strip()

    # If query is already quoted, leave it as is
    if query.startswith('"') and query.endswith('"'):
        return query

    # For natural language queries, don't quote to allow semantic processing
    # Microsoft Graph Search API works better with natural language for semantic understanding
    return query


def request(
    method: str,
    path: str,
    account_id: str | None = None,
    params: dict[str, Any] | None = None,
    json: dict[str, Any] | None = None,
    data: bytes | None = None,
    max_retries: int = 3,
) -> dict[str, Any] | None:
    headers = {
        "Authorization": f"Bearer {get_token(account_id)}",
    }

    if method == "GET":
        if "$search" in (params or {}):
            headers["Prefer"] = 'outlook.body-content-type="text"'
        elif "body" in (params or {}).get("$select", ""):
            headers["Prefer"] = 'outlook.body-content-type="text"'
    else:
        headers["Content-Type"] = (
            "application/json" if json else "application/octet-stream"
        )

    if params and (
        "$search" in params
        or "contains(" in params.get("$filter", "")
        or "/any(" in params.get("$filter", "")
    ):
        headers["ConsistencyLevel"] = "eventual"
        params.setdefault("$count", "true")

    retry_count = 0
    while retry_count <= max_retries:
        try:
            response = _client.request(
                method=method,
                url=f"{BASE_URL}{path}",
                headers=headers,
                params=params,
                json=json,
                content=data,
            )

            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", "5"))
                if retry_count < max_retries:
                    time.sleep(min(retry_after, 60))
                    retry_count += 1
                    continue

            if response.status_code >= 500 and retry_count < max_retries:
                wait_time = (2**retry_count) * 1
                time.sleep(wait_time)
                retry_count += 1
                continue

            response.raise_for_status()

            if response.content:
                return response.json()
            return None

        except httpx.HTTPStatusError as e:
            if retry_count < max_retries and e.response.status_code >= 500:
                wait_time = (2**retry_count) * 1
                time.sleep(wait_time)
                retry_count += 1
                continue
            raise

    return None


def request_paginated(
    path: str,
    account_id: str | None = None,
    params: dict[str, Any] | None = None,
    limit: int | None = None,
) -> Iterator[dict[str, Any]]:
    """Make paginated requests following @odata.nextLink"""
    items_returned = 0
    next_link = None

    while True:
        if next_link:
            result = request("GET", next_link.replace(BASE_URL, ""), account_id)
        else:
            result = request("GET", path, account_id, params=params)

        if not result:
            break

        if "value" in result:
            for item in result["value"]:
                if limit and items_returned >= limit:
                    return
                yield item
                items_returned += 1

        next_link = result.get("@odata.nextLink")
        if not next_link:
            break


def download_raw(
    path: str, account_id: str | None = None, max_retries: int = 3
) -> bytes:
    headers = {"Authorization": f"Bearer {get_token(account_id)}"}

    retry_count = 0
    while retry_count <= max_retries:
        try:
            response = _client.get(f"{BASE_URL}{path}", headers=headers)

            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", "5"))
                if retry_count < max_retries:
                    time.sleep(min(retry_after, 60))
                    retry_count += 1
                    continue

            if response.status_code >= 500 and retry_count < max_retries:
                wait_time = (2**retry_count) * 1
                time.sleep(wait_time)
                retry_count += 1
                continue

            response.raise_for_status()
            return response.content

        except httpx.HTTPStatusError as e:
            if retry_count < max_retries and e.response.status_code >= 500:
                wait_time = (2**retry_count) * 1
                time.sleep(wait_time)
                retry_count += 1
                continue
            raise

    raise ValueError("Failed to download file after all retries")


def _do_chunked_upload(
    upload_url: str,
    data: bytes,
    headers: dict[str, str],
) -> dict[str, Any]:
    """Internal helper for chunked uploads"""
    file_size = len(data)

    for i in range(0, file_size, UPLOAD_CHUNK_SIZE):
        chunk_start = i
        chunk_end = min(i + UPLOAD_CHUNK_SIZE, file_size)
        chunk = data[chunk_start:chunk_end]

        chunk_headers = headers.copy()
        chunk_headers["Content-Length"] = str(len(chunk))
        chunk_headers["Content-Range"] = (
            f"bytes {chunk_start}-{chunk_end - 1}/{file_size}"
        )

        retry_count = 0
        while retry_count <= 3:
            try:
                response = _client.put(upload_url, content=chunk, headers=chunk_headers)

                if response.status_code == 429:
                    retry_after = int(response.headers.get("Retry-After", "5"))
                    if retry_count < 3:
                        time.sleep(min(retry_after, 60))
                        retry_count += 1
                        continue

                response.raise_for_status()

                if response.status_code in (200, 201):
                    return response.json()
                break

            except httpx.HTTPStatusError as e:
                if retry_count < 3 and e.response.status_code >= 500:
                    time.sleep((2**retry_count) * 1)
                    retry_count += 1
                    continue
                raise

    raise ValueError("Upload completed but no final response received")


def create_upload_session(
    path: str,
    account_id: str | None = None,
    item_properties: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Create an upload session for large files"""
    payload = {"item": item_properties or {}}
    result = request("POST", f"{path}/createUploadSession", account_id, json=payload)
    if not result:
        raise ValueError("Failed to create upload session")
    return result


def upload_large_file(
    path: str,
    data: bytes,
    account_id: str | None = None,
    item_properties: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Upload a large file using upload sessions"""
    file_size = len(data)

    if file_size <= UPLOAD_CHUNK_SIZE:
        result = request("PUT", f"{path}/content", account_id, data=data)
        if not result:
            raise ValueError("Failed to upload file")
        return result

    session = create_upload_session(path, account_id, item_properties)
    upload_url = session["uploadUrl"]

    headers = {"Authorization": f"Bearer {get_token(account_id)}"}
    return _do_chunked_upload(upload_url, data, headers)


def create_mail_upload_session(
    message_id: str,
    attachment_item: dict[str, Any],
    account_id: str | None = None,
) -> dict[str, Any]:
    """Create an upload session for large mail attachments"""
    result = request(
        "POST",
        f"/me/messages/{message_id}/attachments/createUploadSession",
        account_id,
        json={"AttachmentItem": attachment_item},
    )
    if not result:
        raise ValueError("Failed to create mail attachment upload session")
    return result


def upload_large_mail_attachment(
    message_id: str,
    name: str,
    data: bytes,
    account_id: str | None = None,
    content_type: str = "application/octet-stream",
) -> dict[str, Any]:
    """Upload a large mail attachment using upload sessions"""
    file_size = len(data)

    attachment_item = {
        "attachmentType": "file",
        "name": name,
        "size": file_size,
        "contentType": content_type,
    }

    session = create_mail_upload_session(message_id, attachment_item, account_id)
    upload_url = session["uploadUrl"]

    headers = {"Authorization": f"Bearer {get_token(account_id)}"}
    return _do_chunked_upload(upload_url, data, headers)


def search_query(
    query: str,
    entity_types: list[str],
    account_id: str | None = None,
    limit: int = 50,
    fields: list[str] | None = None,
    semantic_search: bool = False,
    relevance_threshold: float = 0.0,
) -> Iterator[dict[str, Any]]:
    """Use the modern /search/query API endpoint with optional semantic enhancements"""

    # Process query for semantic search
    if semantic_search:
        # Enhance query for better semantic understanding
        processed_query = _enhance_query_for_semantic_search(query)
    else:
        processed_query = query

    payload = {
        "requests": [
            {
                "entityTypes": entity_types,
                "query": {"queryString": processed_query},
                "size": min(limit, 25),
                "from": 0,
                "includeHiddenContent": True,
            }
        ]
    }

    if fields:
        payload["requests"][0]["fields"] = fields

    # Add semantic search enhancements
    if semantic_search:
        # Enable search result ranking and content analysis
        payload["requests"][0]["enableTopResults"] = True
        payload["requests"][0]["contentSources"] = ["/me"]

    items_returned = 0

    while True:
        result = request("POST", "/search/query", account_id, json=payload)

        if not result or "value" not in result:
            break

        for response in result["value"]:
            if "hitsContainers" in response:
                for container in response["hitsContainers"]:
                    if "hits" in container:
                        for hit in container["hits"]:
                            if limit and items_returned >= limit:
                                return

                            # Apply relevance threshold filtering for semantic search
                            if semantic_search and relevance_threshold > 0.0:
                                hit_score = hit.get("hitScore", 0.0)
                                if hit_score < relevance_threshold:
                                    continue

                            resource = hit["resource"]

                            # Add relevance score to resource for semantic search
                            if semantic_search:
                                resource["_relevance_score"] = hit.get("hitScore", 0.0)
                                resource["_search_rank"] = hit.get(
                                    "rank", items_returned + 1
                                )

                            yield resource
                            items_returned += 1

        if "@odata.nextLink" in result:
            break

        has_more = False
        for response in result.get("value", []):
            for container in response.get("hitsContainers", []):
                if container.get("moreResultsAvailable"):
                    has_more = True
                    break

        if not has_more:
            break

        payload["requests"][0]["from"] += payload["requests"][0]["size"]
