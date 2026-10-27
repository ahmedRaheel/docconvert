# Deploying Pagewise on your EC2 box

Pagewise is one container (Flask + gunicorn + LibreOffice) listening on port 8000 inside Docker.
It sits beside your existing app: different container, different host port.

## 1. Copy and start

```bash
scp -r pagewise ec2-user@YOUR_EC2:~/        # or git clone / zip upload
ssh ec2-user@YOUR_EC2
cd pagewise
docker compose up -d --build
docker compose logs -f pagewise            # wait for "Listening at: http://0.0.0.0:8000"
```

Open `http://YOUR_EC2_IP:8081`. (Allow inbound TCP 8081 in the instance's security group while testing,
then close it once the reverse proxy is in place.)

* Needs about 2 GB of free RAM and 2.5 GB of disk for the image (LibreOffice is large). A t3.small is the minimum; t3.medium is comfortable.
* First build takes several minutes.

## 2. Put it behind your existing reverse proxy (recommended)

### Option A: its own sub-domain (cleanest) `convert.example.com`

```nginx
server {
    listen 80;
    server_name convert.example.com;
    client_max_body_size 60m;               # a bit above MAX_UPLOAD_MB

    location / {
        proxy_pass http://127.0.0.1:8081;   # or http://pagewise:8000 if nginx is in the same Docker network
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 300s;
        proxy_send_timeout 300s;
    }
}
```
Then `sudo certbot --nginx -d convert.example.com` for HTTPS.

### Option B: a path on your existing domain `example.com/convert/`

```nginx
location /convert/ {
    proxy_pass http://127.0.0.1:8081/;
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
    proxy_set_header X-Forwarded-Prefix /convert;   # the app builds all its links from this
    client_max_body_size 60m;
    proxy_read_timeout 300s;
}
```

If your existing nginx runs as a container, either publish nothing and attach Pagewise to the same
Docker network (uncomment the `networks` lines in `docker-compose.yml`, then use `proxy_pass http://pagewise:8000/;`),
or keep the host port and point nginx at the host's private IP.

## 3. Settings (environment variables, or a `.env` file next to docker-compose.yml)

| Variable | Default | Meaning |
|---|---|---|
| `PAGEWISE_PORT` | 8081 | host port |
| `PAGEWISE_BIND` | 0.0.0.0 | set `127.0.0.1` when only a local reverse proxy should reach it |
| `APP_NAME` | Pagewise | name shown in the header and zip file |
| `MAX_UPLOAD_MB` | 50 | per-file limit (keep nginx `client_max_body_size` above it) |
| `RESULT_TTL_SECONDS` | 1800 | how long converted files stay downloadable |
| `MAX_CONCURRENT_CONVERSIONS` | 2 | simultaneous conversions per worker (2 workers) |
| `CONVERSION_TIMEOUT` | 180 | seconds before LibreOffice is stopped |
| `BASIC_AUTH_USER` / `BASIC_AUTH_PASSWORD` | empty | set both to require a login |

## 4. Before it faces the public internet

A converter that accepts any file is a target for abuse. At minimum:

1. Set `BASIC_AUTH_USER` and `BASIC_AUTH_PASSWORD`, or restrict by IP / put it behind your existing login.
2. Add rate limiting in nginx: `limit_req_zone $binary_remote_addr zone=conv:10m rate=10r/m;` and `limit_req zone=conv burst=10;` on `location /api/convert`.
3. Serve it over HTTPS.

## 5. Operating

```bash
docker compose logs -f pagewise      # logs
docker compose up -d --build         # redeploy after changes
docker compose down                  # stop (uploads are already deleted after each conversion)
curl localhost:8081/healthz          # {"status":"ok"}
```
