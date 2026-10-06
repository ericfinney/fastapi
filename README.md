---
title: FastAPI
description: A FastAPI server
tags:
  - fastapi
  - hypercorn
  - python
---

# FastAPI Example

This example starts up a [FastAPI](https://fastapi.tiangolo.com/) server.

[![Deploy on Railway](https://railway.app/button.svg)](https://railway.app/template/-NvLj4?referralCode=CRJ8FE)
## ✨ Features

- FastAPI
- [Hypercorn](https://hypercorn.readthedocs.io/)
- Python 3

## 💁‍♀️ How to use

- Clone locally and install packages with pip using `pip install -r requirements.txt`
- Run locally using `hypercorn main:app --reload`

## 📝 Notes

- To learn about how to use FastAPI with most of its features, you can visit the [FastAPI Documentation](https://fastapi.tiangolo.com/tutorial/)
- To learn about Hypercorn and how to configure it, read their [Documentation](https://hypercorn.readthedocs.io/)

## 🐳 Container / AWS deployment

The `Dockerfile` builds an image that runs on any container host (AWS App Runner,
ECS/Fargate, EC2, Elastic Beanstalk):

```sh
docker build -t boyd-api .
docker run -p 8000:8000 -e PUBLIC_BASE_URL=https://your-domain -e ACTION_API_KEY=... boyd-api
```

| Variable | Purpose |
| --- | --- |
| `PORT` | Listen port (default `8000`) |
| `PUBLIC_BASE_URL` | Origin used in returned `download_url` links. Falls back to `RAILWAY_PUBLIC_URL`, then to the request's own URL |
| `ACTION_API_KEY` | If set, `/actions/generate_proposal_from_pdf` requires `X-API-Key` |
| `BOYD_OUTPUT_DIR` | Where generated workbooks are written (default `/tmp/output`) |

Health check: `GET /health`.

Generated workbooks are stored on the container's local disk and served from
`/download/...`, so run a **single instance** (or add sticky sessions / shared
storage such as S3) — otherwise a download can land on an instance that
doesn't have the file.

The `railway.*`, `nixpacks.toml` and `procfile` files are only used by Railway
and can be deleted once the Railway service is retired.
