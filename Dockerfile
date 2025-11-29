FROM nginx:alpine
COPY index.html styles.css app.js /usr/share/nginx/html/
EXPOSE 80
