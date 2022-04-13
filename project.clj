(defproject genformat "1.0.0"
  ;;  "1.0.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "EPL-2.0 OR GPL-2.0-or-later WITH Classpath-exception-2.0"
            :url "https://www.eclipse.org/legal/epl-2.0/"}
  :dependencies [[org.clojure/data.csv "1.0.0"]
                 [org.clojars.mjdowney/excel-clj "2.0.1"]
                 [org.clojure/clojure "1.10.3"]
                 [aero/aero "1.1.6"]
                 [com.github.seancorfield/next.jdbc "1.2.761"]
                 [org.postgresql/postgresql "42.2.23"]
                 [com.github.seancorfield/honeysql "2.2.858"]]
  :main ^:skip-aot genformat.core
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all
                       :jvm-opts ["-Dclojure.compiler.direct-linking=true"]}})
