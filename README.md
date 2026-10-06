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
- [Uvicorn](https://www.uvicorn.org/)
- Python 3

## 💁‍♀️ How to use

- Clone locally and install packages with pip using `pip install -r requirements.txt`
- Run locally using `uvicorn main:app --reload`

## 📝 Notes

- To learn about how to use FastAPI with most of its features, you can visit the [FastAPI Documentation](https://fastapi.tiangolo.com/tutorial/)
- To learn about Hypercorn and how to configure it, read their [Documentation](https://hypercorn.readthedocs.io/)

## ☁️ AWS deployment (Lambda + S3)

The app runs on AWS Lambda behind a Lambda Function URL (HTTPS, no servers to
manage); generated workbooks are stored in a private S3 bucket and deleted
after 7 days. At low traffic this costs effectively nothing.

How it fits together: `Dockerfile` builds a normal FastAPI container that also
includes the [AWS Lambda Web Adapter](https://github.com/awslabs/aws-lambda-web-adapter),
and `template.yaml` (AWS SAM) creates the function, its URL, the bucket and
permissions. The same image still runs anywhere with `docker run`.

### Prerequisites

- AWS CLI configured for your account (`aws configure` or SSO)
- [AWS SAM CLI](https://docs.aws.amazon.com/serverless-application-model/latest/developerguide/install-sam-cli.html)
- Docker running locally

### First deploy

```sh
sam build
sam deploy --guided
```

Answers for the guided prompts:

| Prompt | Answer |
| --- | --- |
| Stack Name | e.g. `boyd-proposals` |
| AWS Region | your region, e.g. `us-east-1` |
| Parameter ActionApiKey | a long random string (`openssl rand -hex 32`); the GPT Action sends it as `X-API-Key` |
| Parameter PublicBaseUrl | leave empty (uses the Function URL) |
| Parameter FileRetentionDays | `7` |
| Allow SAM CLI IAM role creation | `Y` |
| ProposalFunction Function Url has no authentication. Is this okay? | `Y` (the app checks `X-API-Key` itself) |
| Create managed ECR repositories for all functions? | `Y` |
| Save arguments to configuration file | `Y` |

The `FunctionUrl` output is your new API base URL. Check it:

```sh
curl https://<function-url>/health   # expect "storage": "s3"
```

Later deploys: `sam build && sam deploy`.

### Environment variables

| Variable | Purpose |
| --- | --- |
| `PUBLIC_BASE_URL` | Origin used in returned `download_url` links. Falls back to `RAILWAY_PUBLIC_URL`, then to the request's own URL |
| `ACTION_API_KEY` | If set, `/actions/generate_proposal_from_pdf` requires `X-API-Key` |
| `BOYD_S3_BUCKET` | If set, workbooks are stored in this bucket and `/download/...` redirects to a short-lived signed S3 link. If unset, they're kept on local disk (single-instance hosts only) |
| `BOYD_S3_PREFIX` | Key prefix in the bucket (default `proposals/`) |
| `BOYD_DOWNLOAD_URL_TTL_SECONDS` | Lifetime of each signed S3 link (default `300`) |
| `BOYD_OUTPUT_DIR` | Scratch / local storage directory (default `/tmp/output`) |
| `PORT` | Listen port (default `8000`) |

### Limits on Lambda

- Request bodies are capped at 6 MB, so PDFs uploaded through the web page
  (sent as base64) must be under ~4.5 MB. The GPT Action endpoint is not
  affected — it fetches the PDF from OpenAI's link.
- The first request after a quiet period takes a few extra seconds (cold start).

### Running locally

```sh
docker build -t boyd-api .
docker run -p 8000:8000 boyd-api
```
