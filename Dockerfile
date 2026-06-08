# QA Eval Suite — single container: FastAPI backend that also serves the UI.
# Includes the Azure CLI so it can reuse the host's `az login` (mounted at runtime).
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    DEBIAN_FRONTEND=noninteractive

# Install Azure CLI (needed by AzureCliCredential to mint tokens from az login).
RUN apt-get update && apt-get install -y --no-install-recommends \
        curl ca-certificates apt-transport-https lsb-release gnupg && \
    curl -sL https://aka.ms/InstallAzureCLIDeb | bash && \
    apt-get purge -y curl gnupg && apt-get autoremove -y && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY app ./app

EXPOSE 8080
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8080"]
