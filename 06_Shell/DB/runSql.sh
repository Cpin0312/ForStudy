#!/bin/bash

echo "é¿çsÉpÉX : `dirname ${0}`"

PWD="piadmin"
HOST="10.211.247.104"
PORT=5432
USER=piadmin
DATABASE=sspcpostgre

for filename in `dirname ${0}`/*.sql; do
 	 echo $filename
  
	PGPASSWORD=${PWD} psql -h ${HOST} -p ${PORT} -U ${USER} -d ${DATABASE} -a -f ${filename}
  
done

exit 0