#!/bin/zsh

conda activate mcpo

mcpo --port 8000 --host 0.0.0.0 --config ./config.json --api-key "top-secret"
