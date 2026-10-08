# EduTrack Academic — GitHub + Render deployment

## 1. Create GitHub repository

Create an empty repository named `edutrack-academic`.

In the project folder:

```bat
git init
git add .
git commit -m "Initial EduTrack Academic portal"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/edutrack-academic.git
git push -u origin main
```

## 2. Deploy the web application

GitHub Pages cannot run the Flask backend. Connect the repository to a Python web host such as Render.

Use:

```text
Build command: pip install -r requirements.txt
Start command: gunicorn app:app
```

Set an environment variable:

```text
SCHOOLPULSE_SECRET=<long-random-secret>
```

`render.yaml` is included in this repository.

## 3. Database and uploads for production

The included SQLite database is excellent for a local exhibition. A production school deployment should migrate the database to PostgreSQL and use persistent/object storage for files.

The reason is that normal cloud web instances can be replaced during deployment/restart, while local SQLite/files are stored on the instance filesystem.

## 4. School website integration

`templates/embed.html` provides an integration page. Once the portal is hosted, the school's existing website can link to the portal or embed it where the school's website platform permits iframes.

Example:

```html
<iframe
  src="https://YOUR-PORTAL-DOMAIN.example/embed"
  title="EduTrack Academic"
  style="width:100%;min-height:850px;border:0;border-radius:16px;"
  loading="lazy">
</iframe>
```

Use the hosted HTTPS URL, not `127.0.0.1`.

## 5. Security before real-school use

- Use HTTPS.
- Use PostgreSQL for multi-user production data.
- Use persistent/private object storage for documents.
- Keep secrets in environment variables.
- Add admin approval / teacher identity verification.
- Add password reset and account recovery.
- Add audit logs for marks and attendance changes.
- Back up the database.
- Restrict sensitive documents such as infirmary material.
- Do not put student data in GitHub.
