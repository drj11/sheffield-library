#!/bin/sh

mkdir -p work
for z in source/*.zip
do
    unzip -d work/ "$z"
done
