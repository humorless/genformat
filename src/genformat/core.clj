(ns genformat.core
  (:require [genformat.gen :as gen])
  (:gen-class))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  (dorun (map gen/from-csv args)))
