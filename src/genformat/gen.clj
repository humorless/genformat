(ns genformat.core
  (:require [excel-clj.core :as excel]
            [excel-clj.cell :as cell]))

(def basic-info-style {:border-top :medium
                       :border-bottom :thick
                       :font {:bold true}})

(def basic-info
  [#:excel{:wrapped? true
           :data "S編號"
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

(defn decorate-progress-e [cells]
  (mapv
   (fn [c]
     (assoc c :excel/style {:fill-pattern :solid-foreground
                            :fill-foreground-color [220 220 255]}))
   cells))

(def progress-last-row-style {:border-bottom :thick})

(def progress-value-style {})

(defn decorate-border-c1 [cells]
  (map-indexed
   (fn [idx c]
     (assoc c :excel/style (wrap-five-thick-border idx progress-last-row-style)))
   cells))

(defn decorate-border-not-c1 [cells]
  (map-indexed
   (fn [idx c]
     (assoc c :excel/style (wrap-five-thick-border idx progress-value-style)))
   cells))

(defn decorate-m0 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [c-id (assoc c0 :excel/style {:vertical-alignment :center}
                    :excel/dims {:height 6 :width 1})
        c-name (assoc c1 :excel/style {:vertical-alignment :center}
                      :excel/dims {:height 6 :width 1})
        c-rank (assoc c2 :excel/style {:vertical-alignment :center}
                      :excel/dims {:height 6 :width 1})
        c-status (assoc c3 :excel/style {:vertical-alignment :center}
                        :excel/dims {:height 2 :width 1})
        c-change (assoc c4 :excel/style {:vertical-alignment :center}
                        :excel/dims {:height 2 :width 1})
        c-subject (assoc c5 :excel/style {:vertical-alignment :center}
                         :excel/dims {:height 2 :width 1})]
    [c-id c-name c-rank c-status c-change c-subject c6]))

(defn decorate-m1 [v]
  v)

(defn decorate-e0 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [c-status (assoc c3 :excel/style {:vertical-alignment :center}
                        :excel/dims {:height 2 :width 1})
        c-change (assoc c4 :excel/style {:vertical-alignment :center}
                        :excel/dims {:height 2 :width 1})
        c-subject (assoc c5 :excel/style {:vertical-alignment :center}
                         :excel/dims {:height 2 :width 1})]
    [c0 c1 c2 c-status c-change c-subject c6]))

(defn decorate-e1 [v]
  v)

(defn decorate-c0 [[c0 c1 c2 c3 c4 c5 c6]]
  (let [c-status (assoc c3 :excel/style {:vertical-alignment :center}
                        :excel/dims {:height 2 :width 1})
        c-change (assoc c4 :excel/style {:vertical-alignment :center}
                        :excel/dims {:height 2 :width 1})
        c-subject (assoc c5 :excel/style {:vertical-alignment :center}
                         :excel/dims {:height 2 :width 1})]
    [c0 c1 c2 c-status c-change c-subject c6]))

(defn decorate-c1 [v]
  v)

(defn row-tmpl [subject row-index {:keys [s-id s-name rank status]}]
  (let [basic-cells [#:excel{:wrapped? true, :data s-id}
                     #:excel{:wrapped? true, :data s-name}
                     #:excel{:wrapped? true, :data rank}
                     #:excel{:wrapped? true, :data status}
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
        progress-cells* (cond-> progress-cells-tmpl
                         (and (= subject "E")) decorate-progress-e)
        progress-cells (if (and (= subject "C") (= row-index 1))
                         (decorate-border-c1 progress-cells*)
                         (decorate-border-not-c1 progress-cells*))
        _ (prn progress-cells)]
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

;; Here begin the operational section
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
  (excel/quick-open! workbook))
