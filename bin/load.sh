#!/usr/bin/env bash

PGOPTIONS="--search_path=public"
export PGOPTIONS

psql -d kuops_dev -c "DROP TABLE IF EXISTS gen_subticket;"
psql -d kuops_dev -c "CREATE TABLE gen_subticket (
  path text,
  file_name text,
  student_symbol text,
  student_name text,
  school_grade_name text,
  school_grade_ord integer,
  M text,
  E text,
  C text
);"

psql -d kuops_dev -c "\copy gen_subticket FROM '/Users/laurencechen/Downloads/gen-subticket.csv' HEADER CSV;"
