#!/bin/zsh

conda activate mcpo

mcpo --port 8000 --host 127.0.0.1 --config ./config.json --api-key "top-secret"
