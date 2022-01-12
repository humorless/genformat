#!/usr/bin/env bash

export GENPATH="/Users/laurence/kumon/genformat/target/uberjar/genformat-1.0.0-standalone.jar"

function transform () {
    java -jar $GENPATH $1
}

## 考慮在包含學生資料的 directory 下執行
for file in $(ls .)
do
    ##echo "find the file" $file
    transform $file
done
