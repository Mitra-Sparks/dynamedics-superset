#!/bin/bash

set -eu

EXEC=$(realpath "${0#./}")
EXEC_BASE="${EXEC%/*}"
EXEC_NAME=$(basename "$EXEC")

# defines log function
log() {
  local message="$1"
  printf '%b\n' "$message"
}

# displays usage information
usage() {
  printf %s "Usage: $EXEC_NAME [OPTIONS]

  Provides an interface to control all Superset components.

  [OPTIONS]
    -u|d  - [required] Used to define required action, 'up', 'down'.
" 1>&2
  exit $?
}

NUMARGS=$#
if [ $NUMARGS -le 0 ]; then
  usage
fi

down() {
  echo "Shutting the Superset components down..."
  docker-compose -f $EXEC_BASE/docker-compose-non-dev.yml --env-file $EXEC_BASE/.env down
}

up() {
  echo "Running the Superset components..."
  docker-compose -f $EXEC_BASE/docker-compose-non-dev.yml --env-file $EXEC_BASE/.env up -d
}

UP="" 
DOWN=""

while getopts ":ud" ARG; do
  case "${ARG}" in
    u)  UP=1 ;;
    d)  DOWN=1 ;;
    \?)
        echo "Option not allowed."
        usage
        ;;
    *) usage ;;
  esac
done

shift $((OPTIND-1))  #This tells getopts to move on to the next argument.

if [ -n "${DOWN}" ]; then
    down
elif [ -n "${UP}" ]; then
    up
else
    log "Unknown Execution Option."
fi