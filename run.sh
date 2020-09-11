#!/usr/bin/env bash

set -e

cd ~/projects/NYUCourantInstitute-RoomUpdate/

export DISPLAY=:0

export DISPLAY=localhost:1

pip install -r requirements.txt

python Automator.py

#deactivate
