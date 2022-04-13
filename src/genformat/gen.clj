(ns genformat.gen
  (:require [genformat.sql :as sql]
            [clojure.set :as cset]
            [clojure.data.csv :as csv]
            [clojure.java.io :as io]
            [excel-clj.core :as excel]
            [excel-clj.cell :as cell]))

(def basic-info-style {:border-top :medium
                       :border-bottom :thick
                       :font {:bold true}})

(def basic-info
  [#:excel{:wrapped? true
           :data "學生ID"
           :style basic-info-style}
   #:excel{:wrapped? true
           :data "姓名"
           :style basic-info-style}
   #:excel{:wrapped? true
           :data "學年"
           :style basic-info-style}
   #:excel{:wrapped? true
           :data "狀態"
           :style basic-info-style}
   #:excel{:wrapped? true
           :data "異動"
           :style basic-info-style}
   #:excel{:wrapped? true
           :data "科目"
           :style basic-info-style}
   #:excel{:wrapped? true
           :data "階段"
           :style (assoc basic-info-style
                         :border-right :thick
                         :fill-pattern :solid-foreground
                         :fill-foreground-color [220 220 255])}])

(def progress-info-style {:border-bottom :thick
                          :font {:bold true}
                          :fill-pattern :solid-foreground
                          :fill-foreground-color [220 220 255]})

(defn wrap-five-thick-border
  "every five columns, there is a thick right border"
  [index ori-style]
  (if (= 4 (mod index 5))
    (assoc ori-style :border-right :thick)
    ori-style))

(def progress-info
  (map
   (fn [r]
     #:excel{:wrapped? true
             :data (str
                    (+ 1
                       (* 10 r)))
             :style (wrap-five-thick-border r progress-info-style)})
   (range 20)))

(def header-row
  (into []  (concat
             basic-info
             progress-info)))

(def progress-cells-tmpl
  (map
   (fn [_]
     #:excel{:wrapped? true
             :data ""})
   (range 20)))

(def last-row-style {:border-bottom :thick})

(def last-column-style {:border-right :thick})

(def last-row-last-column-style {:border-bottom :thick
                                 :border-right :thick})

(def empty-style {})

(def vspan-style {:vertical-alignment :center})

(def vspan-last-column-style (merge last-column-style vspan-style))

(def vspan-last-row-style (merge last-row-style vspan-style))

(defn decorate-border-c1 [cells]
  (map-indexed
   (fn [idx c]
     (assoc c :excel/style (wrap-five-thick-border idx last-row-style)))
   cells))

(defn decorate-border-not-c1 [cells]
  (map-indexed
   (fn [idx c]
     (assoc c :excel/style (wrap-five-thick-border idx empty-style)))
   cells))

(defn decorate-progress-e [cells]
  (map-indexed
   (fn [idx c]
     (let [style* (assoc empty-style
                         :fill-pattern :solid-foreground
                         :fill-foreground-color [220 220 255])]
       (assoc c :excel/style (wrap-five-thick-border idx style*))))
   cells))

(defn decorate-m0 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [c-id (assoc c0 :excel/style vspan-style
                    :excel/dims {:height 6 :width 1})
        c-name (assoc c1 :excel/style  vspan-style
                      :excel/dims {:height 6 :width 1})
        c-rank (assoc c2 :excel/style vspan-style
                      :excel/dims {:height 6 :width 1})
        status (get-in c3 [:excel/user :m-status])
        c-status (assoc c3 :excel/style vspan-style
                        :excel/dims {:height 2 :width 1}
                        :excel/data status)
        c-change (assoc c4 :excel/style vspan-style
                        :excel/dims {:height 2 :width 1})
        c-subject (assoc c5 :excel/style vspan-last-column-style
                         :excel/dims {:height 2 :width 1})
        c-stage (assoc c6 :excel/style last-column-style)]
    [c-id c-name c-rank c-status c-change c-subject c-stage]))

(defn decorate-m1 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [status (get-in c3 [:excel/user :m-status])
        c-status (assoc c3 :excel/data status)
        c-subject (assoc c5 :excel/style last-column-style)
        c-stage (assoc c6 :excel/style last-column-style)]
    [c0 c1 c2 c-status c4 c-subject c-stage]))

(defn decorate-e0 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [status (get-in c3 [:excel/user :e-status])
        c-status (assoc c3 :excel/style vspan-style
                        :excel/dims {:height 2 :width 1}
                        :excel/data status)
        c-change (assoc c4 :excel/style vspan-style
                        :excel/dims {:height 2 :width 1})
        c-subject (assoc c5 :excel/style vspan-last-column-style
                         :excel/dims {:height 2 :width 1})
        c-stage  (assoc c6 :excel/style last-column-style)]
    [c0 c1 c2 c-status c-change c-subject c-stage]))

(defn decorate-e1 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [status (get-in c3 [:excel/user :e-status])
        c-status (assoc c3 :excel/data status)
        c-subject (assoc c5 :excel/style last-column-style)
        c-stage (assoc c6 :excel/style last-column-style)]
    [c0 c1 c2 c-status c4 c-subject c-stage]))

(defn decorate-c0 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [status (get-in c3 [:excel/user :c-status])
        c-status (assoc c3 :excel/style vspan-style
                        :excel/dims {:height 2 :width 1}
                        :excel/data status)
        c-change (assoc c4 :excel/style vspan-style
                        :excel/dims {:height 2 :width 1})
        c-subject (assoc c5 :excel/style vspan-last-column-style
                         :excel/dims {:height 2 :width 1})
        c-stage  (assoc c6 :excel/style last-column-style)]
    [c0 c1 c2 c-status c-change c-subject c-stage]))

(defn decorate-c1 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [c-id (assoc c0 :excel/style last-row-style)
        c-name (assoc c1 :excel/style  last-row-style)
        c-rank (assoc c2 :excel/style last-row-style)
        status (get-in c3 [:excel/user :c-status])
        c-status (assoc c3 :excel/style last-row-style
                        :excel/data status)
        c-change (assoc c4 :excel/style last-row-style)
        c-subject (assoc c5 :excel/style last-row-last-column-style)
        c-stage (assoc c6 :excel/style last-row-last-column-style)]
    [c-id c-name c-rank c-status c-change c-subject c-stage]))

(defn row-tmpl [subject row-index {:keys [s-id s-name rank] :as user}]
  (let [basic-cells [#:excel{:wrapped? true, :data s-id}
                     #:excel{:wrapped? true, :data s-name}
                     #:excel{:wrapped? true, :data rank}
                     #:excel{:wrapped? true, :user user}
                     #:excel{:wrapped? true, :data ""}
                     #:excel{:wrapped? true, :data subject}
                     #:excel{:wrapped? true, :data ""}]
        basic-cells (cond-> basic-cells
                      (and (= subject "M") (= row-index 0)) decorate-m0
                      (and (= subject "M") (= row-index 1)) decorate-m1
                      (and (= subject "E") (= row-index 0)) decorate-e0
                      (and (= subject "E") (= row-index 1)) decorate-e1
                      (and (= subject "C") (= row-index 0)) decorate-c0
                      (and (= subject "C") (= row-index 1)) decorate-c1)
        progress-cells* (if (and (= subject "C") (= row-index 1))
                          (decorate-border-c1 progress-cells-tmpl)
                          (decorate-border-not-c1 progress-cells-tmpl))
        progress-cells (cond-> progress-cells*
                         (and (= subject "E")) decorate-progress-e)
        ;; _ (prn basic-cells)
        ]
    (into []
          (concat basic-cells
                  progress-cells))))

(defn user-chunk [user]
  (list
   (row-tmpl "M" 0 user)
   (row-tmpl "M" 1 user)
   (row-tmpl "E" 0 user)
   (row-tmpl "E" 1 user)
   (row-tmpl "C" 0 user)
   (row-tmpl "C" 1 user)))

(defn data->workbook
  "student-data is in the form of
   (
    [centor-id instructor-name student-symbol student-name grade-name grade-ord M E C]
    ...
   )"
  [student-data]
  (let [data-src  (->> student-data
                       (map (fn [[_ _ student-symbol student-name grade-name _ M-status E-status C-status]]
                              (zipmap [:s-id :s-name :rank :m-status :e-status :c-status]
                                      [student-symbol student-name grade-name M-status E-status C-status]))))
        ;; _ (prn data-src)
        data-rows (mapcat user-chunk data-src)
        data (concat (list header-row) data-rows)]
    {"Generated Sheet" data}))

(defn create-out-filepath! [path]
  (let [segments  (clojure.string/split path #"\.")
        front-segments (vec (butlast segments))
        new-segments (conj front-segments "xlsx")]
    (clojure.string/join "." new-segments)))

;; side effect functions
(defn from-tsv
  "filename needs to have the complete file resources path"
  [path]
  (with-open [reader (io/reader path)]
    (rest
     (doall
      (csv/read-csv reader :separator \tab)))))

(defn segment->workbook
  "segemnt is in the form of
   (
    {:student_name 
     :e 
     :m
     :c ... }
    {..}
    {..}
    ...
   )"
  [segment]
  (let [data-src (map #(cset/rename-keys % {:gen_subticket/student_symbol :s-id
                                            :gen_subticket/student_name :s-name
                                            :gen_subticket/school_grade_name :rank
                                            :gen_subticket/E :e-status
                                            :gen_subticket/M :m-status
                                            :gen_subticket/C :c-status}) segment)
        ;; _ (prn data-src)
        data-rows (mapcat user-chunk data-src)
        data (concat (list header-row) data-rows)
        {:gen_subticket/keys [file_name path]} (first segment)]
    {"Generated Sheet" data
     :dir path
     :path (str path file_name ".xlsx")}))

(defn write-workbook! [{:keys [dir path] :as w}]
  (let [workbook (dissoc w :path :dir)]
    (.mkdirs (java.io.File. dir))
    (excel/write! workbook path)))

;; API
(defn from-tsv-to-excel!
  [path]
  (let [students (from-tsv path)
        workbook (data->workbook students)
        out-path (create-out-filepath! path)]
    (excel/write! workbook out-path)))

(defn from-db-to-excel!
  []
  (let [src (sql/get-subticket-src)
        data-segments (partition-by :gen_subticket/file_name src)
        workbooks (map segment->workbook data-segments)]
    (run! write-workbook! workbooks)))
;; REPL session


(comment
  (from-tsv-to-excel! "/Users/laurence/kumon/genformat/resources/enrollment_fc6304201.tsv")

  (def data-from-db
    (list
     {:s-id "455143" :s-name "田中旨夫" :rank "幼大" :status "N"}
     {:s-id "455169" :s-name "Arne" :rank "幼大" :status "N"}
     {:s-id "455123" :s-name "Laurence" :rank "幼大" :status "N"}))

  (def data
    (let [data-rows  (mapcat user-chunk data-from-db)]
      (concat (list header-row)
              data-rows)))

  (let [;; A workbook is any [key value] seq of [sheet-name, sheet-grid].
      ;; Convert the table to a grid with the table-grid function.
        workbook {"My Generated Sheet" data}]
    (excel/quick-open! workbook)))
