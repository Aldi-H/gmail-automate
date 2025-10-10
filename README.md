
# N8N GMAIL AUTOMATE

## What is This Project

This project retrieves email attachments based on the email subject. It is designed to download a single attachment file in either .xlsx or .xls format, using the same document template. The workflow is simple — the user only needs to enter an email subject through an n8n webhook chat. Once all files are saved to the drive, they are automatically merged using a Python script that I’ve already created.

## 
## Setup Steps

1. Clone this repositories

```bash
https://github.com/Aldi-H/gmail-automate.git
cd gmail-automate
```

2. Build docker images
```bash
docker-compose build
```

3. Add execution access to run bash script
```bash
chmod +x add-downloads-path.sh
```

4. Run configuration bash script
```bash
./add-downloads-path.sh
```

5. Run docker container
```bash
docker-compose up -d
```