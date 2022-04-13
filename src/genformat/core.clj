(ns genformat.core
  (:require [genformat.gen :as gen])
  (:gen-class))

(defn -main
  "java -jar xxx.jar $CMD $PARAMETER

  Example:
   java -jar xxx.jar :tsv $the_tsv_input_file_path
   java -jar xxx.jar :db $the_table_name"
  [& args]
  (let [cmd (first args)]
    (cond
      (= cmd ":tsv") (dorun (map gen/from-tsv-to-excel! (rest args)))
      ;; (= cmd ":db") (gen/from-db-to-excel!)
      )))
