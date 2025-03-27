cd bot
docker build -t ssobot .

docker stop ssobot
docker rm ssobot

docker run -d --env-file .env -p 3978:3978 --name ssobot ssobot
