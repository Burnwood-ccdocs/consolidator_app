# Appointment Consolidator App

A Python application that consolidates appointment data from multiple Google Sheets into a single target spreadsheet.

## Overview

This application:
1. Reads a master spreadsheet containing URLs to multiple appointment sheets
2. Extracts data from each appointment sheet
3. Consolidates the data into a single target spreadsheet
4. Uses row hashing to prevent duplicate entries

## Requirements

- Python 3.11+
- Google Sheets API access
- Service account with appropriate permissions

## Dependencies

```
google-api-python-client==2.86.0
google-auth==2.22.0
google-auth-httplib2==0.1.0
google-auth-oauthlib==1.0.0
```

## Setup

1. Clone this repository
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Create a Google Cloud project and enable the Google Sheets API
4. Create a service account and download the JSON key file
5. Rename the key file to `service-account.json` and place it in the project root
6. Grant the service account access to your Google Sheets

## Configuration

Edit the following variables in `consolidate_appointments_all_fields.py`:

- `MASTER_SPREADSHEET_ID`: ID of the spreadsheet containing appointment sheet URLs
- `MASTER_SHEET_NAME`: Name of the tab containing the URLs
- `COMPANY_COLUMN_NAME`: Column header for company names
- `URL_COLUMN_NAME`: Column header for appointment sheet URLs
- `TARGET_SPREADSHEET_ID`: ID of the destination spreadsheet
- `TARGET_SHEET_NAME`: Name of the tab to write data to

## Usage

Run the script:

```
python consolidate_appointments_all_fields.py
```

The script will:
1. Read all appointment sheet URLs from the master spreadsheet
2. Process each sheet and extract appointment data
3. Consolidate the data into the target spreadsheet
4. Track processed rows to avoid duplicates

## Docker Deployment

### Build the Docker image

```
docker build -t consolidator-app .
```

### Run the container

```
docker run -it --rm consolidator-app
```

### Deployment Options

#### Local Docker

```
docker run -d --name consolidator-app consolidator-app
```

#### Google Cloud Run

```
# Tag your image
docker tag consolidator-app gcr.io/[PROJECT-ID]/consolidator-app

# Push to Google Container Registry
docker push gcr.io/[PROJECT-ID]/consolidator-app

# Deploy to Cloud Run
gcloud run deploy consolidator-app --image gcr.io/[PROJECT-ID]/consolidator-app --platform managed
```

#### Azure Container Instances

```
# Tag your image
docker tag consolidator-app [YOUR-REGISTRY].azurecr.io/consolidator-app

# Push to Azure Container Registry
docker push [YOUR-REGISTRY].azurecr.io/consolidator-app

# Deploy to ACI
az container create --resource-group [RESOURCE-GROUP] --name consolidator-app --image [YOUR-REGISTRY].azurecr.io/consolidator-app
```

#### AWS ECS/Fargate

```
# Tag your image
docker tag consolidator-app [AWS-ACCOUNT-ID].dkr.ecr.[REGION].amazonaws.com/consolidator-app

# Push to ECR
aws ecr get-login-password | docker login --username AWS --password-stdin [AWS-ACCOUNT-ID].dkr.ecr.[REGION].amazonaws.com
docker push [AWS-ACCOUNT-ID].dkr.ecr.[REGION].amazonaws.com/consolidator-app
```

## Logs

To view logs from a running Docker container:

```
docker logs consolidator-app
```

For cloud deployments, check the respective platform's logging interface.

## Data Persistence

The application uses `processed_rows.json` to track which rows have been processed to avoid duplicates. This file is automatically created and updated during execution. 
