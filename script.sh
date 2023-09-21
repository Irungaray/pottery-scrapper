# !/bin/bash

TODAY=$(date +%d-%m-%Y)

[ ! -e  "./html" ] && mkdir ./html
[ ! -e  "./html/$TODAY" ] && mkdir ./html/$TODAY
[ ! -e  "./json" ] && mkdir ./json
[ ! -e  "./json/$TODAY" ] && mkdir ./json/$TODAY
[ ! -e  "./xlsx" ] && mkdir ./xlsx

curl https://dpcolors.com/productos1/oxidos/?mpage=3 >> ./html/$TODAY/oxidos.html
curl https://dpcolors.com/productos1/pigmentos-puros/?mpage=5 >> ./html/$TODAY/pigmentos-puros.html
curl https://dpcolors.com/productos1/esmaltes-ceramicos/?mpage=9 >> ./html/$TODAY/esmaltes-ceramicos.html

node ./script.js

[ -e  "./xlsx/$TODAY.xlsx" ] && xdg-open ./xlsx/$TODAY.xlsx
