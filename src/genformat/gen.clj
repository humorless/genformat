(ns genformat.core
  (:require [excel-clj.core :as excel]
            [excel-clj.cell :as cell]))

(def basic-info
  [#:excel{:wrapped? true
           :data "S編號"
           :style {:border-bottom :thin, :font {:bold true}}}
   #:excel{:wrapped? true
           :data "姓名"
           :style {:border-bottom :thin, :font {:bold true}}}
   #:excel{:wrapped? true
           :data "學年"
           :style {:border-bottom :thin, :font {:bold true}}}
   #:excel{:wrapped? true
           :data "狀態"
           :style {:border-bottom :thin, :font {:bold true}}}
   #:excel{:wrapped? true
           :data "異動"
           :style {:border-bottom :thin, :font {:bold true}}}
   #:excel{:wrapped? true
           :data "科目"
           :style {:border-bottom :thin, :font {:bold true}}}
   #:excel{:wrapped? true
           :data "階段"
           :style {:border-bottom :thin, :font {:bold true}}}])

(def progress-v (mapv
                 (fn [r]
                   #:excel{:wrapped? true
                           :data (str
                                  (+ 1
                                     (* 10 r)))
                           :style {:border-bottom :thin, :font {:bold true}}})
                 (range 20)))

(def header
  (into []  (concat
             basic-info
             progress-v)))

(def data
  (list
   header
   [#:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "abcde", :style {:vertical-alignment :center} :dims {:height 6 :width 1}}
    #:excel{:wrapped? true, :data "幼大", :style {:vertical-alignment :center} :dims {:height 6 :width 1}}
    #:excel{:wrapped? true, :data "N", :style {:vertical-alignment :center} :dims {:height 2 :width 1}}
    #:excel{:wrapped? true, :data "", :style {:vertical-alignment :center} :dims {:height 2 :width 1}}
    #:excel{:wrapped? true, :data "M", :style {:vertical-alignment :center} :dims {:height 2 :width 1}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {:fill-pattern :solid-foreground
                                              :fill-foreground-color [220 220 255]}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}]
   [#:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "abcde", :style {}}
    #:excel{:wrapped? true, :data "幼大", :style {}}
    #:excel{:wrapped? true, :data "N", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "M", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}]
   [#:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "abcde", :style {}}
    #:excel{:wrapped? true, :data "幼大", :style {}}
    #:excel{:wrapped? true, :data "N", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "E", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}]
   [#:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "abcde", :style {}}
    #:excel{:wrapped? true, :data "幼大", :style {}}
    #:excel{:wrapped? true, :data "N", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "E", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}]
   [#:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "abcde", :style {}}
    #:excel{:wrapped? true, :data "幼大", :style {}}
    #:excel{:wrapped? true, :data "N", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "C", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}]
   [#:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "abcde", :style {}}
    #:excel{:wrapped? true, :data "幼大", :style {}}
    #:excel{:wrapped? true, :data "N", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "C", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}
    #:excel{:wrapped? true, :data "", :style {}}]))

(let [;; A workbook is any [key value] seq of [sheet-name, sheet-grid].
      ;; Convert the table to a grid with the table-grid function.
      workbook {"My Generated Sheet" data}]
   (excel/quick-open! workbook))
