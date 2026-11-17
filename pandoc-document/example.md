---
title: "Neural Signal Analysis API"
subtitle: "Integration & Operations Guide"
doctype: "Software Documentation"
author:
  - "Sandi Baressi Šegota"
project: "Sample Project"
version: "1.4.0"
date: "June 2026"
toc: true
toc-depth: 2
numbersections: true
lang: en
eu-logo: "eu-logo.png"
# confidential: true   # uncomment to stamp "Confidential" in the footer
---

# Overview

The Neural Signal Analysis API exposes the analysis pipeline over a REST
interface. It ingests time-resolved fluorescence
recordings and returns calcium-signaling, reactive-oxygen-species, and
vesicular-trafficking feature vectors suitable for downstream ensemble models.

This guide covers authentication, the core endpoints, error handling, and
operational notes for self-hosted deployments.

::: note
All examples assume API **v1**. The base URL for the managed service is
`https://api.daignostics.info/v1`.
:::

## Audience

This document is written for backend engineers and integrators who already
have a workspace and an issued API key.

# Authentication

Every request must carry a bearer token in the `Authorization` header. Tokens
are scoped per workspace and can be rotated from the console.

```bash
curl https://api.daignostics.info/v1/recordings \
  -H "Authorization: Bearer $DG_TOKEN" \
  -H "Content-Type: application/json"
```

::: {.warning title="Rotate leaked keys immediately"}
A token grants full read/write access to its workspace. If a key is exposed,
revoke it from **Settings → API Keys** before doing anything else.
:::

# Endpoints

The table below summarizes the most frequently used routes.

| Method | Path                  | Purpose                          |
|:-------|:----------------------|:---------------------------------|
| `GET`  | `/recordings`         | List recordings in the workspace |
| `POST` | `/recordings`         | Upload a new recording           |
| `GET`  | `/recordings/{id}`    | Fetch metadata for one recording |
| `POST` | `/analyses`           | Queue an analysis job            |
| `GET`  | `/analyses/{id}`      | Poll job status and results      |

## Submitting an analysis

Send the recording identifier and a model profile. The call returns a job
handle you can poll until the status becomes `completed`.

```python
import requests

resp = requests.post(
    "https://api.daignostics.info/v1/analyses",
    headers={"Authorization": f"Bearer {token}"},
    json={"recording_id": rec_id, "profile": "ensemble-bayesian"},
)
resp.raise_for_status()
job = resp.json()
print(job["id"], job["status"])
```

::: tip
Polling at a 2–5 s interval is plenty; jobs are typically scheduled within a
second and most complete in under a minute.
:::

# Error handling

Errors use standard HTTP status codes and a JSON body with a stable
machine-readable `code`:

1. `400` — malformed request payload
2. `401` — missing or invalid token
3. `404` — resource not found in this workspace
4. `429` — rate limit exceeded; honor the `Retry-After` header
5. `500` — internal error; safe to retry with backoff

# Operational notes

For self-hosted deployments, the worker pool reads from a Redis queue and
writes feature vectors to object storage. Scale workers horizontally to match
ingest volume.

::: important
The processing workers require GPU access. Running the ensemble profile on
CPU-only nodes will fall back to a slower path and may exceed request timeouts.
:::

That's the whole surface area — see the console for live request logs and
per-key usage.
