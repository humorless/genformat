(ns genformat.sql
  (:require [honey.sql :as sql]
            [next.jdbc :as jdbc]
            [aero.core :as aero]))

(def config (aero/read-config "config.edn"))

(def db-entry {:jdbcUrl
               (:database-url config)})

(def conn (jdbc/get-datasource db-entry))

(def subticket-map
  {:select [:*]
   :from [:gen_subticket]})

(defn get-subticket-src []
  (jdbc/execute! conn (sql/format subticket-map)))
