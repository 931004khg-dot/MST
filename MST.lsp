;;; ========================================================================
;;; MST (수량집계표작성) - 통합 버전
;;; Part 1~4 통합 완료 (줌 기능 및 괄호 파싱 포함)
;;; ========================================================================

(vl-load-com)

;; VBA 상수 정의 (Table 서식에 필요)
(setq acTitleRow  1)
(setq acHeaderRow 2)
(setq acDataRow   4)
(setq acAlignmentTopRight 3)
(setq acAlignMiddleRight 6)
(setq acAlignmentBottomRight 9)

;; 전역 변수 초기화
(setq *mst-temp-group-count* 0)
(setq *mst-temp-numbers* nil)
(setq *mst-temp-notes* nil)


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; A. 최하위 유틸리티 및 보조 함수
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun mst-ensure-layers-on (/ doc layers layer-names layer-obj)
  (princ "\n>> 필수 레이어를 활성화합니다...")
  (setq layer-names '("!-MST-TEMP"))
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  (setq layers (vla-get-layers doc))
  (foreach layer-name layer-names
    (if (tblsearch "LAYER" layer-name)
      (vl-catch-all-apply
        '(lambda ()
           (setq layer-obj (vla-item layers layer-name))
           (vla-put-LayerOn layer-obj :vlax-true)
           (vla-put-Freeze layer-obj :vlax-false)
         )
      )
    )
  )
  (princ " 완료.")
)

(defun mst-get-topmost-text-position (text-ent-list / top-y top-text-center-x current-y vla-obj min-pt-sa max-pt-sa min-pt max-pt)
  (if text-ent-list
    (progn
      (setq top-y nil top-text-center-x nil)
      (foreach ent text-ent-list
        (setq vla-obj (vlax-ename->vla-object ent))
        (vla-getboundingbox vla-obj 'min-pt-sa 'max-pt-sa)
        (setq min-pt (vlax-safearray->list min-pt-sa))
        (setq max-pt (vlax-safearray->list max-pt-sa))
        (setq current-y (cadr max-pt))
        (if (or (not top-y) (> current-y top-y))
          (progn
            (setq top-y current-y)
            (setq top-text-center-x (/ (+ (car min-pt) (car max-pt)) 2.0))
          )
        )
      )
      (list top-text-center-x (+ top-y (* text-height 0.5) (* 5.0 (/ text-height 3.5))) 0.0)
    )
    '(0 0 0)
  )
)

(defun mst-get-lwpolyline-points (input-pline)
  (mapcar 'cdr (vl-remove-if-not '(lambda (x) (= (car x) 10)) (entget input-pline))))

(defun mst-group-exists-p (name / doc dictionaries group-dict)
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  (setq dictionaries (vla-get-dictionaries doc))
  (if (not (vl-catch-all-error-p (setq group-dict (vl-catch-all-apply 'vla-item (list dictionaries "ACAD_GROUP")))))
    (if (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-item (list group-dict name))))
      t
      nil
    )
    nil
  )
)

(defun mst-cleanup (/ region group-name ss-test group-dict group-obj i ent temp-num-ent temp-note-ent)
  (princ "\n>> 정리 작업을 수행합니다...")
  (if temp-regions (foreach region temp-regions (if (and region (entget region)) (entdel region))))
  (setq temp-regions nil)
  (if (and created-group-names (not (vl-catch-all-error-p (setq group-dict (vla-item (vla-get-dictionaries (vla-get-activedocument (vlax-get-acad-object))) "ACAD_GROUP")))))
    (foreach group-name created-group-names
      (if (not (vl-catch-all-error-p (setq group-obj (vl-catch-all-apply 'vla-item (list group-dict group-name))))) (vla-delete group-obj))
    )
  )
  (setq created-group-names nil)
  (if (tblsearch "LAYER" "!-MST-TEMP")
    (if (setq ss-test (ssget "_X" '((8 . "!-MST-TEMP"))))
      (progn (setq i 0) (repeat (sslength ss-test) (setq ent (ssname ss-test i)) (if (and ent (entget ent)) (entdel ent)) (setq i (1+ i)))))
  )
  (setq processed-objects nil)
  (setq *mst-temp-numbers* nil)
  (setq *mst-temp-notes* nil)
  (setq *mst-temp-group-count* 0)
  (setq original-colors nil)
  (princ " 완료.")
  (princ)
)

(defun mst-compare-items (a b / len-a len-b)
  (setq len-a (strlen (nth 1 a))) (setq len-b (strlen (nth 1 b)))
  (cond ((> len-a len-b) t) ((< len-a len-b) nil) ((< (nth 1 a) (nth 1 b)) t) ((> (nth 1 a) (nth 1 b)) nil) ((< (nth 2 a) (nth 2 b)) t) ((> (nth 2 a) (nth 2 b)) nil) (t (< (nth 3 a) (nth 3 b))))
)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; B. 객체 선택, 검증, 임시 영역 생성
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun mst-get-text-centerpoint (ent / vla-obj min-pt-sa max-pt-sa min-pt max-pt)
  (setq vla-obj (vlax-ename->vla-object ent))
  (vla-getboundingbox vla-obj 'min-pt-sa 'max-pt-sa)
  (setq min-pt (vlax-safearray->list min-pt-sa)
        max-pt (vlax-safearray->list max-pt-sa))
  (list (/ (+ (car min-pt) (car max-pt)) 2.0) (/ (+ (cadr min-pt) (cadr max-pt)) 2.0) (/ (+ (caddr min-pt) (caddr max-pt)) 2.0))
)

(defun mst-select-objects ( / choice obj-type target-layer doc current-space ss plines texts leaders tables i ent original_obj coords pt1 pt2 new_pline_vla new_pline_ent err-check layers-col new-layer first-table)
  (princ "\n[1/7] 작업 모드를 결정합니다...")
  
  (setq choice nil)
  (while (not choice)
    (setq choice (entsel "\n>> 수량 데이터 산출 객체(레이어 확인) 또는 수량집계표(총괄수량집계표 생성)를 선택하세요: "))
    (if (not choice) (princ "\n!! 선택하지 않았습니다. 다시 선택해주세요."))
  )
  
  (setq ent (car choice))
  (setq obj-type (cdr (assoc 0 (entget ent))))

  (if (= obj-type "ACAD_TABLE")
    (progn
      (princ "\n>> 수량집계표가 감지되었습니다. 총괄수량집계표 모드로 진입합니다.")
      (setq first-table ent)
      
      (princ "\n>> 집계할 수량집계표를 모두 선택하세요: ") 
      
      (setvar "NOMUTT" 1)
      (princ "\n수량집계표(TABLE) 선택: ")
      (setq ss (vl-catch-all-apply 'ssget (list (list (cons 0 "ACAD_TABLE")))))
      (setvar "NOMUTT" 0)
      
      (if (vl-catch-all-error-p ss) (setq ss nil))
      
      (setq tables (list first-table))
      (if ss
        (progn
          (setq i 0)
          (repeat (sslength ss)
            (setq ent (ssname ss i))
            (if (not (equal ent first-table)) 
                (setq tables (cons ent tables))
            )
            (setq i (1+ i))
          )
        )
      )
      
      (princ (strcat "\n>> 총 " (itoa (length tables)) "개의 수량집계표가 취합되었습니다."))
      (list "TABLE_MODE" tables nil)
    )

    (progn
      (setq target-layer (cdr (assoc 8 (entget ent))))
      (princ (strcat "\n>> '" target-layer "' 레이어를 기준으로 작업을 시작합니다."))
      
      (princ "\n>> 폴리선, 텍스트, 지시선을 선택하세요: ")
      (setq ss (ssget (list (cons 0 "LWPOLYLINE,TEXT,LEADER") (cons 8 target-layer))))
      
      (if ss
        (progn
          (setq doc (vla-get-activedocument (vlax-get-acad-object)))
          (setq current-space (if (= (getvar "CTAB") "Model") (vla-get-modelspace doc) (vla-get-paperspace doc)))
          (setq plines nil texts nil leaders nil i 0)
          
          (repeat (sslength ss)
            (setq ent (ssname ss i)
                  obj-type (cdr (assoc 0 (entget ent))))
            (cond
              ((= obj-type "LWPOLYLINE") (setq plines (cons ent plines)))
              ((= obj-type "TEXT") (setq texts (cons ent texts)))
              ((= obj-type "LEADER") (setq leaders (cons ent leaders)))
            )
            (setq i (1+ i))
          )

          (if leaders
            (progn
              (if (not (tblsearch "LAYER" "!-MST-TEMP"))
                (progn
                  (setq layers-col (vla-get-layers doc))
                  (setq new-layer (vla-add layers-col "!-MST-TEMP"))
                  (vla-put-color new-layer 1)
                  (princ " (임시 레이어 생성됨)")
                )
              )
              (princ (strcat "\n  - 지시선 " (itoa (length leaders)) "개 감지. 변환 작업을 수행합니다..."))
              (foreach leader-ent leaders
                (setq err-check 
                  (vl-catch-all-apply 
                    '(lambda ()
                       (setq original_obj (vlax-ename->vla-object leader-ent))
                       (setq coords (vlax-safearray->list (vlax-variant-value (vla-get-coordinates original_obj))))
                       (if (>= (length coords) 6)
                         (progn
                           (setq pt1 (list (nth (- (length coords) 6) coords) (nth (- (length coords) 5) coords)))
                           (setq pt2 (list (nth (- (length coords) 3) coords) (nth (- (length coords) 2) coords)))
                           (setq new_pline_vla (vla-addlightweightpolyline current-space 
                                                 (vlax-safearray-fill 
                                                   (vlax-make-safearray vlax-vbdouble '(0 . 3)) 
                                                   (list (car pt1) (cadr pt1) (car pt2) (cadr pt2))
                                                 )
                                               )
                           )
                           (vla-put-layer new_pline_vla "!-MST-TEMP")
                           (vla-put-color new_pline_vla 2)
                           (setq new_pline_ent (vlax-vla-object->ename new_pline_vla))
                           (setq plines (cons new_pline_ent plines))
                         )
                       )
                     )
                  )
                )
              )
            )
          )
          
          (princ (strcat "\n>> 총 " (itoa (sslength ss)) "개 객체 처리 완료."))
          
          (if (and plines texts)
            (progn
              (princ (strcat "\n  - 폴리선(지시선 포함): " (itoa (length plines)) "개, 텍스트: " (itoa (length texts)) "개"))
              (list (reverse plines) (reverse texts) target-layer)
            )
            (progn
              (princ "\n*** 오류: 작업에 필요한 폴리선과 텍스트가 모두 발견되지 않았습니다. ***")
              nil
            )
          )
        )
        (progn (princ (strcat "\n>> '" target-layer "' 레이어에서 선택된 객체가 없습니다.")) nil)
      )
    )
  )
)

(defun mst-validate-text-height (texts / first-height)
  (princ "\n[2/7] 텍스트 높이를 검증하고 스케일을 계산합니다...")
  (setq first-height (cdr (assoc 40 (entget (car texts)))))
  (foreach txt (cdr texts) (if (not (equal first-height (cdr (assoc 40 (entget txt))) 1e-6)) (progn (princ (strcat "\n*** 오류: 텍스트 높이가 다릅니다. 기준(" (rtos first-height) "), 오류 객체: " (cdr (assoc 1 (entget txt))))) (exit))))
  (setq text-height first-height scale-factor (/ text-height 3.5) upper-offset (* 14.0 scale-factor) lower-offset (* -7.0 scale-factor))
  (princ (strcat " 완료. (높이: " (rtos text-height 2 2) ", 스케일: " (rtos scale-factor 2 2) ")")) t
)

(defun mst-create-temp-regions (plines / region-list current-pline-entity old-osmode old-cmdecho vla-pline actual-upper-dist actual-lower-dist test-offset-obj-list test-offset-list vla-test-offset min-pt-sa max-pt-sa min-pt-sa-test max-pt-sa-test orig-y test-y offset-results upper-offset-list lower-offset-list vla-upper-offset-pline vla-lower-offset-pline upper-offset-pline-ent lower-offset-pline-ent upper-pts lower-pts region-pts region-ent)
  (princ "\n[3/7] 임시 영역을 생성합니다... ")
  (setq old-cmdecho (getvar "CMDECHO") old-osmode (getvar "OSMODE"))
  (setvar "CMDECHO" 0) (setvar "OSMODE" 0)
  (if (not (tblsearch "LAYER" "!-MST-TEMP")) (command "_.LAYER" "_M" "!-MST-TEMP" "_C" "1" "" ""))
  (setq region-list nil)
  (foreach current-pline-entity plines
    (setq vla-pline (vlax-ename->vla-object current-pline-entity)) (setq actual-upper-dist nil actual-lower-dist nil)
    (setq test-offset-obj-list (vl-catch-all-apply 'vla-offset (list vla-pline 1.0)))
    (if (and (not (vl-catch-all-error-p test-offset-obj-list)) (= 'VARIANT (type test-offset-obj-list)))
      (progn
        (setq test-offset-list (vlax-safearray->list (vlax-variant-value test-offset-obj-list)))
        (if (and (> (length test-offset-list) 0) (= 'VLA-OBJECT (type (setq vla-test-offset (car test-offset-list)))))
          (progn (vla-getBoundingBox vla-pline 'min-pt-sa 'max-pt-sa) (vla-getBoundingBox vla-test-offset 'min-pt-sa-test 'max-pt-sa-test) (setq orig-y (/ (+ (cadr (vlax-safearray->list min-pt-sa)) (cadr (vlax-safearray->list max-pt-sa))) 2.0)) (setq test-y (/ (+ (cadr (vlax-safearray->list min-pt-sa-test)) (cadr (vlax-safearray->list max-pt-sa-test))) 2.0)) (vla-delete vla-test-offset) (if (> test-y orig-y) (setq actual-upper-dist upper-offset actual-lower-dist lower-offset) (setq actual-upper-dist (- upper-offset) actual-lower-dist (- lower-offset))))
        )
      )
    )
    (if (and actual-upper-dist actual-lower-dist)
      (progn
        (setq offset-results (list (vl-catch-all-apply 'vla-offset (list vla-pline actual-upper-dist)) (vl-catch-all-apply 'vla-offset (list vla-pline actual-lower-dist))))
        (if (and (not (vl-catch-all-error-p (car offset-results))) (not (vl-catch-all-error-p (cadr offset-results))) (= (type (car offset-results)) 'VARIANT) (= (type (cadr offset-results)) 'VARIANT))
          (progn
            (setq upper-offset-list (vlax-safearray->list (vlax-variant-value (car offset-results)))) (setq lower-offset-list (vlax-safearray->list (vlax-variant-value (cadr offset-results))))
            (if (and (> (length upper-offset-list) 0) (> (length lower-offset-list) 0) (= 'VLA-OBJECT (type (setq vla-upper-offset-pline (car upper-offset-list)))) (= 'VLA-OBJECT (type (setq vla-lower-offset-pline (car lower-offset-list)))))
              (progn (vla-put-layer vla-upper-offset-pline "!-MST-TEMP") (vla-put-layer vla-lower-offset-pline "!-MST-TEMP") (setq upper-offset-pline-ent (vlax-vla-object->ename vla-upper-offset-pline)) (setq lower-offset-pline-ent (vlax-vla-object->ename vla-lower-offset-pline)) (setq upper-pts (mst-get-lwpolyline-points upper-offset-pline-ent)) (setq lower-pts (mst-get-lwpolyline-points lower-offset-pline-ent)) (if (and upper-pts lower-pts (> (distance (car upper-pts) (car lower-pts)) (distance (car upper-pts) (last lower-pts)))) (setq lower-pts (reverse lower-pts))) (setq region-pts (append upper-pts (reverse lower-pts))) (setq region-ent (entmakex (append '((0 . "LWPOLYLINE") (100 . "AcDbEntity") (8 . "!-MST-TEMP") (100 . "AcDbPolyline")) (list (cons 90 (length region-pts)) '(70 . 1)) (mapcar '(lambda (p) (cons 10 p)) region-pts)))) (if region-ent (progn (vla-put-color (vlax-ename->vla-object region-ent) 2) (setq region-list (cons region-ent region-list)) (princ ".")) (princ "E")) (entdel upper-offset-pline-ent) (entdel lower-offset-pline-ent))
            )
          )
        )
      )
    )
  )
  (setvar "OSMODE" old-osmode) (setvar "CMDECHO" old-cmdecho)
  (setq temp-regions (reverse region-list))
  (if temp-regions (progn (princ " 완료.") t) (progn (princ "\n*** 경고: 유효한 임시 영역을 생성하지 못했습니다. ***") nil))
)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; C. 그룹 찾기 및 색상 지정
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun mst-find-and-color-groups (data target-layer / plines valid-groups region-texts-ss i region-texts region-idx main-pline-entity region-pts txt region-vla-obj min-pt-sa max-pt-sa region-center text-distances sorted-distances sorted-texts group-name group-ss old-cmdecho copied-pline-vla copied-pline-ent copied-texts-ent copied-txt-vla txt-ent txt-vla all-texts-list auto-grouped-texts-list ungrouped-texts-list auto-processed-objects local-created-group-names group-color temp-text-obj note-text note-pos note-obj note-align-pos this-group-notes group-members-sa groups new-group)
  (princ "\n[4/7] 유효한 그룹을 찾고 복사본을 생성합니다...")
  (setq plines (car data) all-texts-list (cadr data) valid-groups nil region-idx 0 local-created-group-names nil auto-grouped-texts-list nil auto-processed-objects nil)
  (setq old-cmdecho (getvar "CMDECHO")) (setvar "CMDECHO" 0)
  (foreach region temp-regions
    (setq main-pline-entity (nth region-idx plines) region-vla-obj (vlax-ename->vla-object region) region-pts (mst-get-lwpolyline-points region) region-texts nil)
    (vla-getBoundingBox region-vla-obj 'min-pt-sa 'max-pt-sa)
    (setq region-center (list (/ (+ (car (vlax-safearray->list min-pt-sa)) (car (vlax-safearray->list max-pt-sa))) 2.0) (/ (+ (cadr (vlax-safearray->list min-pt-sa)) (cadr (vlax-safearray->list max-pt-sa))) 2.0) 0.0))
    (if (setq region-texts-ss (ssget "_CP" region-pts (list '(0 . "TEXT") (cons 8 target-layer))))
      (progn
        (setq i 0 text-distances nil)
        (repeat (sslength region-texts-ss) (setq txt (ssname region-texts-ss i)) (setq text-distances (cons (list (distance region-center (mst-get-text-centerpoint txt)) txt) text-distances)) (setq i (1+ i)))
        (setq text-distances (vl-remove-if '(lambda (x) (member (cadr x) auto-grouped-texts-list)) text-distances))
        (if text-distances
          (progn
            (setq sorted-distances (vl-sort text-distances '(lambda (a b) (< (car a) (car b))))) (setq sorted-texts (mapcar 'cadr sorted-distances)) (setq region-texts nil)
            (if (> (length sorted-texts) 0) (setq region-texts (cons (car sorted-texts) region-texts))) (if (> (length sorted-texts) 1) (setq region-texts (cons (cadr sorted-texts) region-texts)))
            (if (and (> (length sorted-texts) 2) (< (/ (car (nth 2 sorted-distances)) (car (nth 1 sorted-distances))) 3.0)) (setq region-texts (cons (caddr sorted-texts) region-texts)))
            (setq region-texts (reverse region-texts))
            (if (and region-texts (or (= (length region-texts) 2) (= (length region-texts) 3)))
              (progn
                (setq *mst-temp-group-count* (1+ *mst-temp-group-count*))
                (setq copied-pline-vla (vla-copy (vlax-ename->vla-object main-pline-entity))) (setq copied-pline-ent (vlax-vla-object->ename copied-pline-vla)) (setq copied-texts-ent nil this-group-notes nil)
                (setq region-texts (vl-sort region-texts '(lambda (a b) (> (cadr (mst-get-text-centerpoint a)) (cadr (mst-get-text-centerpoint b))))))
                (if (= (length region-texts) 3)
                  (progn
                    (setq txt-ent (nth 0 region-texts) copied-txt-vla (vla-copy (vlax-ename->vla-object txt-ent)) copied-texts-ent (cons (vlax-vla-object->ename copied-txt-vla) copied-texts-ent))
                    (setq note-text "공종" note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq this-group-notes (cons (vlax-vla-object->ename note-obj) this-group-notes))
                    (setq txt-ent (nth 1 region-texts) copied-txt-vla (vla-copy (vlax-ename->vla-object txt-ent)) copied-texts-ent (cons (vlax-vla-object->ename copied-txt-vla) copied-texts-ent))
                    (setq note-text "규격" note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq this-group-notes (cons (vlax-vla-object->ename note-obj) this-group-notes))
                    (setq txt-ent (nth 2 region-texts) copied-txt-vla (vla-copy (vlax-ename->vla-object txt-ent)) copied-texts-ent (cons (vlax-vla-object->ename copied-txt-vla) copied-texts-ent))
                    (setq note-text "수량" note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq this-group-notes (cons (vlax-vla-object->ename note-obj) this-group-notes))
                  )
                  (progn
                    (setq txt-ent (nth 0 region-texts) copied-txt-vla (vla-copy (vlax-ename->vla-object txt-ent)) copied-texts-ent (cons (vlax-vla-object->ename copied-txt-vla) copied-texts-ent))
                    (setq note-text "공종" note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq this-group-notes (cons (vlax-vla-object->ename note-obj) this-group-notes))
                    (setq txt-ent (nth 1 region-texts) copied-txt-vla (vla-copy (vlax-ename->vla-object txt-ent)) copied-texts-ent (cons (vlax-vla-object->ename copied-txt-vla) copied-texts-ent))
                    (setq note-text "수량" note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq this-group-notes (cons (vlax-vla-object->ename note-obj) this-group-notes))
                  )
                )
                (setq *mst-temp-notes* (cons (reverse this-group-notes) *mst-temp-notes*))
                (setq copied-texts-ent (reverse copied-texts-ent)) (setq group-color (if (= (length region-texts) 2) 6 3)) (vla-put-Layer copied-pline-vla "!-MST-TEMP") (vla-put-Color copied-pline-vla group-color) (foreach txt-ent copied-texts-ent (setq txt-vla (vlax-ename->vla-object txt-ent)) (vla-put-Layer txt-vla "!-MST-TEMP") (vla-put-Color txt-vla group-color))
                (setq group-name (strcat "MST_Group_" (itoa (+ (length created-group-names) (length local-created-group-names))))) (if (mst-group-exists-p group-name) (setq group-name (strcat group-name "_" (itoa (fix (getvar "CDATE")))))) (setq group-ss (ssadd)) (foreach txt-ent copied-texts-ent (ssadd txt-ent group-ss)) 

(if (> (sslength group-ss) 0)
  (progn
    (setq group-members-sa (vlax-make-safearray vlax-vbobject (cons 0 (1- (sslength group-ss)))))
    (setq i 0)
    (repeat (sslength group-ss)
      (vlax-safearray-put-element group-members-sa i (vlax-ename->vla-object (ssname group-ss i)))
      (setq i (1+ i))
    )
    (setq groups (vla-get-groups (vla-get-activedocument (vlax-get-acad-object))))
    (setq new-group (vla-add groups group-name))
    (vla-appenditems new-group group-members-sa)
  )
)
                (setq local-created-group-names (cons group-name local-created-group-names)) (setq valid-groups t) (setq auto-grouped-texts-list (append auto-grouped-texts-list region-texts)) (setq auto-processed-objects (cons copied-texts-ent auto-processed-objects))
                (setq temp-text-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) (itoa *mst-temp-group-count*) (vlax-3d-point (mst-get-topmost-text-position region-texts)) text-height)) (vla-put-layer temp-text-obj "!-MST-TEMP") (vla-put-Color temp-text-obj 2) (vla-put-Alignment temp-text-obj 8) (vla-put-TextAlignmentPoint temp-text-obj (vlax-3d-point (mst-get-topmost-text-position region-texts))) (setq *mst-temp-numbers* (cons (vlax-vla-object->ename temp-text-obj) *mst-temp-numbers*))
              )
            )
          )
        )
      )
    )
    (setq region-idx (1+ region-idx))
  )
  (setq ungrouped-texts-list (vl-remove-if '(lambda (x) (member x auto-grouped-texts-list)) all-texts-list))
  (foreach txt ungrouped-texts-list (setq copied-txt-vla (vla-copy (vlax-ename->vla-object txt))) (vla-put-Layer copied-txt-vla "!-MST-TEMP") (vla-put-Color copied-txt-vla 4))
  (setvar "CMDECHO" old-cmdecho)
  
  (setq *mst-temp-numbers* (reverse *mst-temp-numbers*))
  (setq *mst-temp-notes* (reverse *mst-temp-notes*))
  
  (if valid-groups (princ (strcat " " (itoa (length local-created-group-names)) "개 그룹 발견.")))
  (list (reverse local-created-group-names) ungrouped-texts-list (reverse auto-processed-objects))
)

(defun mst-get-group-data-by-member (member-ent all-groups / found-group)
  (setq found-group nil)
  (foreach group all-groups
    (if (and (not found-group) (member member-ent group))
      (setq found-group group)
    )
  )
  found-group
)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; D. 수동 그룹 조정
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun mst-manual-interaction (p-objects c-groups u-texts target-layer force-default / user-opt sel-ent ent-data old-group-data work-type-ent spec-ent qty-ent ss selected-texts-ents new-group-list group-ss group-color group-name copied-vla copied-ent temp-text-obj cyan-copy-ss group-index num-to-del notes-to-del new-note-ents note-text note-pos note-align-pos note-obj sorted-new-group new-num-ent user-confirm recovered-texts orig-pos orig-txt search-ss k found-orig tmp-ent prompt-msg minpt maxpt txt-vla minPoint maxPoint curr-min curr-max dx dy margin-x margin-y)
  (defun get-index (item lst / i found)
    (setq i 0 found nil) (while (and (not found) (< i (length lst))) (if (equal item (nth i lst)) (setq found i)) (setq i (1+ i))) found)
  (defun remove-nth (n lst / i result)
    (setq i 0 result nil) (foreach item lst (if (/= i n) (setq result (cons item result))) (setq i (1+ i))) (reverse result))
  
  (princ "\n\n[5/7] 그룹을 수동으로 조정합니다...")
  
  (initget "Add Revise Finish")
  
  (if (= force-default "Revise")
    (setq prompt-msg "\n>> 수동 그룹 [추가(A)/수정(R)/완료(F)] <수정>: ")
    (if (and u-texts (> (length u-texts) 0))
      (setq prompt-msg "\n>> 수동 그룹 [추가(A)/수정(R)/완료(F)] <추가>: ")
      (setq prompt-msg "\n>> 수동 그룹 [추가(A)/수정(R)/완료(F)] <완료>: ")
    )
  )
  
  (setq user-opt (getkword prompt-msg))

  (if (not user-opt)
    (if (= force-default "Revise")
      (setq user-opt "Revise")
      (if (and u-texts (> (length u-texts) 0))
        (setq user-opt "Add")
        (setq user-opt "Finish")
      )
    )
  )

  (cond
    ((= user-opt "Revise")
      (setq sel-ent (entsel "\n>> 수정할 그룹의 텍스트를 선택하세요: "))
      (if (and sel-ent (setq ent-data (entget (car sel-ent))) (= "!-MST-TEMP" (cdr (assoc 8 ent-data))))
        (progn
          (setq old-group-data (mst-get-group-data-by-member (car sel-ent) p-objects))
          (if old-group-data
            (progn
              (princ "\n>> 기존 그룹을 해제합니다. (텍스트는 유지됨)")
              (setq group-index (get-index old-group-data p-objects))
              (if group-index
                (progn
                  (setq num-to-del (nth group-index *mst-temp-numbers*)) (if (and num-to-del (entget num-to-del)) (entdel num-to-del))
                  (setq *mst-temp-numbers* (remove-nth group-index *mst-temp-numbers*))
                  (setq notes-to-del (nth group-index *mst-temp-notes*)) (foreach n notes-to-del (if (and n (entget n)) (entdel n)))
                  (setq *mst-temp-notes* (remove-nth group-index *mst-temp-notes*))
                )
              )

              (setq recovered-texts nil)
              (foreach item old-group-data 
                (if (and item (entget item)) 
                  (progn
                    (vla-put-Color (vlax-ename->vla-object item) 4)
                    (setq orig-pos (cdr (assoc 10 (entget item))))
                    (setq orig-txt (cdr (assoc 1 (entget item))))
                    (setq search-ss (ssget "_X" (list '(0 . "TEXT") (cons 8 target-layer) (cons 1 orig-txt))))
                    (if search-ss
                      (progn
                        (setq k 0 found-orig nil)
                        (repeat (sslength search-ss)
                           (setq tmp-ent (ssname search-ss k))
                           (if (< (distance (cdr (assoc 10 (entget tmp-ent))) orig-pos) 0.01)
                              (setq found-orig tmp-ent)
                           )
                           (setq k (1+ k))
                        )
                        (if found-orig (setq recovered-texts (cons found-orig recovered-texts)))
                      )
                    )
                  )
                )
              )
              
              (setq p-objects (vl-remove old-group-data p-objects))
              (setq u-texts (append u-texts recovered-texts))

              (setq work-type-ent nil spec-ent nil qty-ent nil)
              (princ "\n   >> 새 [공종] 텍스트를 선택하세요: ") (setq ss (ssget "_+.:S:E" (list '(0 . "TEXT") (cons 8 target-layer))))
              (if (and ss (= 1 (sslength ss)))
                (progn
                  (setq work-type-ent (ssname ss 0)) (princ "\n   >> 새 [규격/수량] 텍스트를 선택하세요: ") (setq ss (ssget "_+.:S:E" (list '(0 . "TEXT") (cons 8 target-layer))))
                  (if (and ss (= 1 (sslength ss)))
                    (progn (setq spec-ent (ssname ss 0)) (initget " ") (princ "\n   >> 새 [수량] 텍스트를 선택하거나 Enter를 누르세요: ") (setq ss (ssget "_+.:S:E" (list '(0 . "TEXT") (cons 8 target-layer)))) (if (and ss (= 1 (sslength ss))) (setq qty-ent (ssname ss 0)) (progn (setq qty-ent spec-ent) (setq spec-ent nil))))
                  )
                )
              )
              
              (if (and work-type-ent qty-ent)
                (progn
                  (setq *mst-temp-group-count* (1+ *mst-temp-group-count*))
                  (setq selected-texts-ents (if spec-ent (list work-type-ent spec-ent qty-ent) (list work-type-ent qty-ent)))
                  (setq group-name (strcat "MST_Manual_Group_" (itoa (getvar "MILLISECS")))) (if (mst-group-exists-p group-name) (setq group-name (strcat group-name "R")))
                  (setq new-group-list nil group-ss (ssadd)) (setq group-color (if (= (length selected-texts-ents) 2) 6 3))

                  (foreach ent selected-texts-ents
                    (setq copied-vla (vla-copy (vlax-ename->vla-object ent))) (setq copied-ent (vlax-vla-object->ename copied-vla)) (vla-put-Layer copied-vla "!-MST-TEMP") (vla-put-Color copied-vla group-color) (ssadd copied-ent group-ss) (setq new-group-list (cons copied-ent new-group-list))
                    (if (setq cyan-copy-ss (ssget "_X" (list '(8 . "!-MST-TEMP") '(62 . 4) (cons 1 (cdr (assoc 1 (entget ent))))))) (vlax-for item (vla-get-ActiveSelectionSet (vla-get-ActiveDocument (vlax-get-acad-object))) (vla-delete item)))
                  )
                  (setq new-group-list (reverse new-group-list)) (command "_.GROUP" "_C" group-name "" group-ss "") (princ "\n>> 그룹 수정 완료!")

                  (setq new-note-ents nil) (setq sorted-new-group (vl-sort new-group-list '(lambda (a b) (> (cadr (mst-get-text-centerpoint a)) (cadr (mst-get-text-centerpoint b))))))
                  (if (= (length sorted-new-group) 3)
                    (progn
                      (setq txt-ent (nth 0 sorted-new-group)) (setq note-text "공종") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents))
                      (setq txt-ent (nth 1 sorted-new-group)) (setq note-text "규격") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents))
                      (setq txt-ent (nth 2 sorted-new-group)) (setq note-text "수량") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents)))
                    (progn
                      (setq txt-ent (nth 0 sorted-new-group)) (setq note-text "공종") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents))
                      (setq txt-ent (nth 1 sorted-new-group)) (setq note-text "수량") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents)))
                  )
                  (setq temp-text-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) (itoa *mst-temp-group-count*) (vlax-3d-point (mst-get-topmost-text-position new-group-list)) text-height))
                  (vla-put-layer temp-text-obj "!-MST-TEMP") (vla-put-Color temp-text-obj 2) (vla-put-Alignment temp-text-obj 8) (vla-put-TextAlignmentPoint temp-text-obj (vlax-3d-point (mst-get-topmost-text-position new-group-list)))
                  (setq new-num-ent (vlax-vla-object->ename temp-text-obj))
                  (setq *mst-temp-numbers* (cons new-num-ent *mst-temp-numbers*)) (setq *mst-temp-notes* (cons (reverse new-note-ents) *mst-temp-notes*))
                  
                  (setq u-texts (vl-remove-if '(lambda (x) (member x selected-texts-ents)) u-texts))
                  (list "CONTINUE" (cons new-group-list p-objects) (cons group-name c-groups) u-texts)
                )
                (progn (princ "\n>> 선택이 취소되어 수정 작업이 중단됩니다.") (list "CONTINUE" p-objects c-groups u-texts))
              )
            )
            (progn (princ "\n!! 선택한 객체에 해당하는 그룹 데이터를 찾지 못했습니다.") (list "CONTINUE" p-objects c-groups u-texts))
          )
        )
        (progn (princ "\n!! 유효한 그룹 객체가 아닙니다. 색상으로 표시된 임시 텍스트를 선택하세요.") (list "CONTINUE" p-objects c-groups u-texts))
      )
    )
    ((= user-opt "Add")
      (if (or (null u-texts) (= (length u-texts) 0))
        (progn
          (princ "\n>> 추가할 객체가 없습니다.")
          (list "CONTINUE" p-objects c-groups u-texts)
        )
        (progn
          (princ (strcat "\n>> 새 그룹 추가 (" (itoa (length u-texts)) "개 미지정 텍스트 남음):"))
          (setq work-type-ent nil spec-ent nil qty-ent nil)

          (while (null work-type-ent)
            (princ "\n   >> [공종] 텍스트를 선택하세요: ")
            (setq ss (ssget "_+.:S:E" (list '(0 . "TEXT") (cons 8 target-layer))))
            (if (and ss (= 1 (sslength ss)))
              (if (not (member (ssname ss 0) u-texts))
                (progn 
                  (princ "\n!! 잘못된 객체")
                  (setq work-type-ent "BACK_TO_MENU")
                )
                (setq work-type-ent (ssname ss 0))
              )
              (progn (princ "\n선택이 취소되었습니다.") (setq work-type-ent "QUIT"))
            )
          )

          (if (= work-type-ent "BACK_TO_MENU")
            (setq spec-ent "BACK_TO_MENU" qty-ent "BACK_TO_MENU")
            (if (and work-type-ent (/= work-type-ent "QUIT"))
              (progn
                (while (null spec-ent)
                  (princ "\n   >> [규격/수량] 텍스트를 선택하세요: ")
                  (setq ss (ssget "_+.:S:E" (list '(0 . "TEXT") (cons 8 target-layer))))
                  (if (and ss (= 1 (sslength ss)))
                    (if (not (member (ssname ss 0) u-texts))
                      (progn 
                        (princ "\n!! 잘못된 객체")
                        (setq spec-ent "BACK_TO_MENU")
                      )
                      (setq spec-ent (ssname ss 0))
                    )
                    (progn (princ "\n선택이 취소되었습니다.") (setq spec-ent "QUIT"))
                  )
                )
              )
            )
          )

          (if (and spec-ent (/= spec-ent "QUIT") (/= spec-ent "BACK_TO_MENU"))
            (progn
              (initget " ") (princ "\n   >> [수량] 텍스트를 선택하거나 Enter를 누르세요: ")
              (setq ss (ssget "_+.:S:E" (list '(0 . "TEXT") (cons 8 target-layer))))
              (if (and ss (= 1 (sslength ss)))
                (if (not (member (ssname ss 0) u-texts))
                  (progn 
                    (princ "\n!! 잘못된 객체")
                    (setq qty-ent "BACK_TO_MENU")
                  )
                  (setq qty-ent (ssname ss 0))
                )
                (progn (setq qty-ent spec-ent) (setq spec-ent nil))
              )
            )
          )
          
          (if (or (= work-type-ent "BACK_TO_MENU") (= spec-ent "BACK_TO_MENU") (= qty-ent "BACK_TO_MENU"))
            (list "REVISE_DEFAULT" p-objects c-groups u-texts)
            (if (and work-type-ent qty-ent (/= work-type-ent "QUIT") (/= spec-ent "QUIT") (/= qty-ent "INVALID"))
              (progn
              (setq *mst-temp-group-count* (1+ *mst-temp-group-count*))
              (setq selected-texts-ents (if spec-ent (list work-type-ent spec-ent qty-ent) (list work-type-ent qty-ent)))
              (setq group-name (strcat "MST_Manual_Group_" (itoa (getvar "MILLISECS")))) (if (mst-group-exists-p group-name) (setq group-name (strcat group-name "R")))
              (setq new-group-list nil group-ss (ssadd)) (setq group-color (if (= (length selected-texts-ents) 2) 6 3))
              (foreach ent selected-texts-ents
                (setq copied-vla (vla-copy (vlax-ename->vla-object ent))) (setq copied-ent (vlax-vla-object->ename copied-vla)) (vla-put-Layer copied-vla "!-MST-TEMP") (vla-put-Color copied-vla group-color) (ssadd copied-ent group-ss) (setq new-group-list (cons copied-ent new-group-list))
                (if (setq cyan-copy-ss (ssget "_X" (list '(8 . "!-MST-TEMP") '(62 . 4) (cons 1 (cdr (assoc 1 (entget ent))))))) (vlax-for item (vla-get-ActiveSelectionSet (vla-get-ActiveDocument (vlax-get-acad-object))) (vla-delete item)))
              )
              (setq new-group-list (reverse new-group-list)) (command "_.GROUP" "_C" group-name "" group-ss "") (princ "\n>> 새 수동 그룹 생성 완료!")

              (setq new-note-ents nil) (setq sorted-new-group (vl-sort new-group-list '(lambda (a b) (> (cadr (mst-get-text-centerpoint a)) (cadr (mst-get-text-centerpoint b))))))
              (if (= (length sorted-new-group) 3)
                (progn
                  (setq txt-ent (nth 0 sorted-new-group)) (setq note-text "공종") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents))
                  (setq txt-ent (nth 1 sorted-new-group)) (setq note-text "규격") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents))
                  (setq txt-ent (nth 2 sorted-new-group)) (setq note-text "수량") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents)))
                (progn
                  (setq txt-ent (nth 0 sorted-new-group)) (setq note-text "공종") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents))
                  (setq txt-ent (nth 1 sorted-new-group)) (setq note-text "수량") (setq note-pos (cdr (assoc 10 (entget txt-ent)))) (setq note-align-pos (list (- (car note-pos) text-height) (cadr note-pos) (caddr note-pos))) (setq note-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) note-text (vlax-3d-point note-pos) (/ text-height 3.0))) (vla-put-Alignment note-obj acAlignmentBottomRight) (vla-put-TextAlignmentPoint note-obj (vlax-3d-point note-align-pos)) (vla-put-layer note-obj "!-MST-TEMP") (setq new-note-ents (cons (vlax-vla-object->ename note-obj) new-note-ents)))
              )
              (setq temp-text-obj (vla-addtext (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-paperspace (vla-get-activedocument (vlax-get-acad-object)))) (itoa *mst-temp-group-count*) (vlax-3d-point (mst-get-topmost-text-position new-group-list)) text-height))
              (vla-put-layer temp-text-obj "!-MST-TEMP") (vla-put-Color temp-text-obj 2) (vla-put-Alignment temp-text-obj 8) (vla-put-TextAlignmentPoint temp-text-obj (vlax-3d-point (mst-get-topmost-text-position new-group-list)))
              (setq new-num-ent (vlax-vla-object->ename temp-text-obj))
              (setq *mst-temp-numbers* (cons new-num-ent *mst-temp-numbers*)) (setq *mst-temp-notes* (cons (reverse new-note-ents) *mst-temp-notes*))

              (list "CONTINUE" (cons new-group-list p-objects) (cons group-name c-groups) (vl-remove-if '(lambda (x) (member x selected-texts-ents)) u-texts))
              )
              (list "FINISH" p-objects c-groups u-texts)
            )
          )
        )
      )
    )
    ((= user-opt "Finish")
      (if (and u-texts (> (length u-texts) 0))
        (progn
          ;; Zoom to ungrouped objects to help user locate them
          (if (> (length u-texts) 0)
            (progn
              (setq minpt nil maxpt nil)
              ;; Calculate bounding box for all ungrouped texts
              (foreach txt-ent u-texts
                (if (and txt-ent (entget txt-ent))
                  (progn
                    (setq txt-vla (vlax-ename->vla-object txt-ent))
                    (vla-getBoundingBox txt-vla 'minPoint 'maxPoint)
                    (setq curr-min (vlax-safearray->list minPoint))
                    (setq curr-max (vlax-safearray->list maxPoint))
                    (if (not minpt)
                      (setq minpt curr-min maxpt curr-max)
                      (progn
                        (setq minpt (list (min (car minpt) (car curr-min))
                                         (min (cadr minpt) (cadr curr-min))
                                         (min (caddr minpt) (caddr curr-min))))
                        (setq maxpt (list (max (car maxpt) (car curr-max))
                                         (max (cadr maxpt) (cadr curr-max))
                                         (max (caddr maxpt) (caddr curr-max))))
                      )
                    )
                  )
                )
              )
              ;; Add margin around bounding box (20% on each side)
              (if (and minpt maxpt)
                (progn
                  (setq dx (- (car maxpt) (car minpt)))
                  (setq dy (- (cadr maxpt) (cadr minpt)))
                  (setq margin-x (* dx 0.2))
                  (setq margin-y (* dy 0.2))
                  (setq minpt (list (- (car minpt) margin-x) (- (cadr minpt) margin-y) (caddr minpt)))
                  (setq maxpt (list (+ (car maxpt) margin-x) (+ (cadr maxpt) margin-y) (caddr maxpt)))
                  ;; Zoom to the calculated window
                  (command "_.ZOOM" "_W" minpt maxpt)
                )
              )
            )
          )
          (initget "Y N")
          (setq user-confirm (getkword "\n>> 경고: 아직 그룹화되지 않은 객체가 남아있습니다. 그래도 진행하시겠습니까? [Yes/No] <No>: "))
          (if (or (not user-confirm) (= user-confirm "N"))
            (progn
              (princ "\n>> 작업을 계속합니다.")
              (list "CONTINUE" p-objects c-groups u-texts)
            )
            (list "FINISH" p-objects c-groups u-texts)
          )
        )
        (list "FINISH" p-objects c-groups u-texts)
      )
    )
  )
)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; E. 데이터 분석 및 테이블 생성
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun mst-analyze-groups (/ result group texts work-type-txt spec-txt qty-txt work-type-str spec-str qty-str qty-data qty-num qty-unit combined-key existing continue-flag group-number user-resp proceed-with-add minpt maxpt txt-vla minPoint maxPoint curr-min curr-max dx dy margin-x margin-y)
  (princ "\n[6/7] 데이터를 분석하고 집계합니다...")
  (setq result nil continue-flag t group-number 0) 
  (foreach group processed-objects 
    (if continue-flag
      (progn
        (setq group-number (1+ group-number))
        (setq work-type-txt nil spec-txt nil qty-txt nil proceed-with-add t) 

        (setq texts (vl-sort group '(lambda (a b) (> (cadr (mst-get-text-centerpoint a)) (cadr (mst-get-text-centerpoint b))))))
        
        (cond
          ((= (length texts) 3)
           (setq work-type-txt (nth 0 texts))
           (setq spec-txt (nth 1 texts))
           (setq qty-txt (nth 2 texts)))
          ((= (length texts) 2)
           (setq work-type-txt (nth 0 texts))
           (setq qty-txt (nth 1 texts)))
        )

        (if qty-txt
          (progn
            (setq work-type-str (if work-type-txt (cdr (assoc 1 (entget work-type-txt))) "공종없음"))
            (setq spec-str (if spec-txt (cdr (assoc 1 (entget spec-txt))) ""))
            (setq qty-str (cdr (assoc 1 (entget qty-txt))))
            (setq qty-data (mst-parse-quantity qty-str))

            (if (and spec-txt (/= "" spec-str) (not (and (wcmatch spec-str "(*") (wcmatch spec-str "*)"))))
              (progn
                ;; Zoom to the problematic group to help user locate it
                (setq minpt nil maxpt nil)
                ;; Calculate bounding box for all texts in this group
                (foreach txt-ent texts
                  (if (and txt-ent (entget txt-ent))
                    (progn
                      (setq txt-vla (vlax-ename->vla-object txt-ent))
                      (vla-getBoundingBox txt-vla 'minPoint 'maxPoint)
                      (setq curr-min (vlax-safearray->list minPoint))
                      (setq curr-max (vlax-safearray->list maxPoint))
                      (if (not minpt)
                        (setq minpt curr-min maxpt curr-max)
                        (progn
                          (setq minpt (list (min (car minpt) (car curr-min))
                                           (min (cadr minpt) (cadr curr-min))
                                           (min (caddr minpt) (caddr curr-min))))
                          (setq maxpt (list (max (car maxpt) (car curr-max))
                                           (max (cadr maxpt) (cadr curr-max))
                                           (max (caddr maxpt) (caddr curr-max))))
                        )
                      )
                    )
                  )
                )
                ;; Add margin around bounding box (20% on each side)
                (if (and minpt maxpt)
                  (progn
                    (setq dx (- (car maxpt) (car minpt)))
                    (setq dy (- (cadr maxpt) (cadr minpt)))
                    (setq margin-x (* dx 0.2))
                    (setq margin-y (* dy 0.2))
                    (setq minpt (list (- (car minpt) margin-x) (- (cadr minpt) margin-y) (caddr minpt)))
                    (setq maxpt (list (+ (car maxpt) margin-x) (+ (cadr maxpt) margin-y) (caddr maxpt)))
                    ;; Zoom to the calculated window
                    (command "_.ZOOM" "_W" minpt maxpt)
                  )
                )
                
                (princ (strcat "\n*** 오류: 그룹 #" (itoa group-number) "의 규격 텍스트에 괄호가 없습니다."))
                (initget "Y N 무시 중단")
                (setq user-resp (getkword (strcat "\n>> 그룹 #" (itoa group-number) "에 괄호가 없지만 집계에 포함(무시)하시겠습니까? [Yes(무시)/No(중단)] <Y>: ")))
                
                (if (or (not user-resp) (= user-resp "Y") (= user-resp "무시"))
                  (progn
                    (princ (strcat "\n   (경고: 그룹 #" (itoa group-number) " 괄호 없음. 집계에 포함합니다.)"))
                    (setq proceed-with-add t)
                  )
                  (progn
                    (princ "\n>> 사용자가 중단을 선택했습니다. 확인 후 다시 실행하세요.")
                    (setq continue-flag nil)
                    (setq proceed-with-add nil)
                  )
                )
              )
              (setq proceed-with-add t)
            )

            (if (and continue-flag proceed-with-add)
              (progn
                (setq spec-str (if (and (wcmatch spec-str "(*") (wcmatch spec-str "*)")) (substr spec-str 2 (- (strlen spec-str) 2)) spec-str))
                (if qty-data 
                  (progn 
                    (setq qty-num (car qty-data)) 
                    (setq qty-unit (cadr qty-data)) 
                    (setq combined-key (strcat work-type-str "|" spec-str "|" qty-unit)) 
                    (setq existing (assoc combined-key result)) 
                    (if existing 
                      (setq result (subst (list combined-key work-type-str spec-str qty-unit (+ (nth 4 existing) qty-num)) existing result)) 
                      (setq result (cons (list combined-key work-type-str spec-str qty-unit qty-num) result))
                    )
                  ) 
                  (princ (strcat "\n경고: 수량 분석 실패 - '" qty-str "'"))
                )
              )
            )
          )
          (princ (strcat "\n경고: 그룹 #" (itoa group-number) "의 텍스트 구성 오류. 건너뜁니다."))
        )
      )
    )
  )
  (if continue-flag 
    (progn 
      (princ (strcat " " (itoa (length result)) "개 항목으로 집계 완료.")) 
      result 
    )
    nil 
  )
)

(defun mst-parse-quantity (qty-string / clean-str eq-pos i len char num-str unit-str split-pos)
  (setq clean-str (vl-string-trim " " qty-string)) (if (setq eq-pos (vl-string-search "=" clean-str)) (setq clean-str (vl-string-trim " " (substr clean-str (+ eq-pos 2))))) (setq clean-str (vl-string-subst "" "," clean-str)) (setq len (strlen clean-str) i 1 split-pos nil) (while (and (<= i len) (not split-pos)) (setq char (substr clean-str i 1)) (if (not (wcmatch char "[0-9.]")) (setq split-pos i)) (setq i (1+ i))) (if split-pos (if (= split-pos 1) (list 0.0 clean-str) (progn (setq num-str (substr clean-str 1 (1- split-pos))) (setq unit-str (vl-string-trim " " (substr clean-str split-pos))) (list (atof num-str) unit-str))) (list (atof clean-str) ""))
)

(defun mst-parse-work-spec (work-str spec-str / open-pos close-pos inside-text outside-text final-work final-spec)
  ;; 공종 텍스트에서 괄호가 있으면 괄호 안 내용을 규격으로, 괄호 밖 내용을 공종으로 분리
  ;; work-str: 공종 텍스트, spec-str: 기존 규격 텍스트
  ;; 반환값: (공종텍스트 . 규격텍스트)
  (if (and work-str (setq open-pos (vl-string-search "(" work-str)))
    (progn
      (setq close-pos (vl-string-search ")" work-str))
      (if (and close-pos (> close-pos open-pos))
        (progn
          ;; 괄호 안의 텍스트 추출 (괄호 제외)
          (setq inside-text (vl-string-trim " " (substr work-str (+ open-pos 2) (- close-pos open-pos 1))))
          ;; 괄호 밖의 텍스트 추출
          (setq outside-text (vl-string-trim " " 
            (strcat 
              (substr work-str 1 open-pos)
              (if (< close-pos (strlen work-str))
                (substr work-str (+ close-pos 2))
                ""
              )
            )
          ))
          ;; 규격이 비어있으면 괄호 안 내용을 규격으로
          (if (or (not spec-str) (= spec-str ""))
            (cons outside-text inside-text)
            (cons work-str spec-str)
          )
        )
        ;; 여는 괄호만 있고 닫는 괄호가 없으면 원본 유지
        (cons work-str spec-str)
      )
    )
    ;; 괄호가 없으면 원본 유지
    (cons work-str spec-str)
  )
)

(defun mst-aggregate-tables (table-list / data-list vla-tbl row-count r work spec unit qty qty-val key existing)
  (princ "\n>> 선택된 수량집계표의 데이터를 종합합니다...")
  (setq data-list nil)
  (foreach ent table-list
    (setq vla-tbl (vlax-ename->vla-object ent))
    (setq row-count (vla-get-rows vla-tbl))
    (setq r 2)
    (while (< r row-count)
      (setq work (vla-gettext vla-tbl r 0))
      (setq spec (vla-gettext vla-tbl r 1))
      (setq unit (vla-gettext vla-tbl r 2))
      (setq qty (vla-gettext vla-tbl r 3))
      
      (if (and work spec unit qty (/= qty ""))
        (progn
          (setq qty-val (atof (vl-string-subst "" "," qty)))
          (setq key (strcat work "|" spec "|" unit))
          
          (setq existing (assoc key data-list))
          (if existing
            (setq data-list (subst (list key work spec unit (+ (nth 4 existing) qty-val)) existing data-list))
            (setq data-list (cons (list key work spec unit qty-val) data-list))
          )
        )
      )
      (setq r (1+ r))
    )
  )
  data-list
)

(defun mst-sort-data (data) (vl-sort data 'mst-compare-items))

(defun mst-create-table (data text-style title-text / sorted-data table-pos table-obj i row-data col-widths title-height col block-space qty-val qty-str old-osmode parsed-work-spec final-work final-spec)
  (princ (strcat "\n[최종] " title-text "를 생성합니다..."))
  (if (and data (> (length data) 0))
    (progn
      (setq sorted-data (mst-sort-data data))
      (initget 1) (setq table-pos (getpoint "\n>> 표를 삽입할 위치를 지정하세요: "))
      (if table-pos
        (progn
          (setq old-osmode (getvar "OSMODE")) (setvar "OSMODE" 0)
          (setq block-space (if (= (getvar "TILEMODE") 1) (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (vla-get-block (vla-get-activelayout (vla-get-activedocument (vlax-get-acad-object))))))
          (if (not (tblsearch "LAYER" "!-수량표")) (command "_.LAYER" "_M" "!-수량표" "_C" "7" ""))
          (setvar "OSMODE" old-osmode)
          (setq table-obj (vla-addtable block-space (vlax-3d-point table-pos) (+ 2 (length sorted-data)) 5 (* text-height (/ 10.0 3.0)) 1.0))
          
          (vla-put-Layer table-obj "!-수량표") 
          (vla-SetTextStyle table-obj acTitleRow text-style) (vla-SetTextStyle table-obj acHeaderRow text-style) (vla-SetTextStyle table-obj acDataRow text-style) 
          (vla-SetTextHeight table-obj acHeaderRow text-height) (vla-SetTextHeight table-obj acDataRow text-height) 
          
          (vla-put-horzcellmargin table-obj (* text-height 0.5)) 
          (vla-put-vertcellmargin table-obj (* text-height 0.1))

          (setq col-widths (list (* 64.0 (/ text-height 3.5)) (* 60.0 (/ text-height 3.5)) (* 30.0 (/ text-height 3.5)) (* 50.0 (/ text-height 3.5)) (* 30.0 (/ text-height 3.5))))
          (setq i 0) (foreach width col-widths (vla-SetColumnWidth table-obj i width) (setq i (1+ i)))
          
          (vla-settext table-obj 0 0 title-text) 
          (setq title-height (* text-height (/ 6.0 3.5))) (vla-SetCellTextHeight table-obj 0 0 title-height) (vla-mergecells table-obj 0 0 0 4) (vla-SetCellAlignment table-obj 0 0 5)
          (vla-settext table-obj 1 0 "공 종") (vla-settext table-obj 1 1 "규 격") (vla-settext table-obj 1 2 "단 위") (vla-settext table-obj 1 3 "수 량") (vla-settext table-obj 1 4 "비 고")
          (setq col 0) (while (< col 5) (vla-SetCellAlignment table-obj 1 col 5) (setq col (1+ col)))
          (setq row-height-target (* text-height (/ 10.0 3.0))) 
          (vla-SetRowHeight table-obj 0 row-height-target)
          (vla-SetRowHeight table-obj 1 row-height-target)
          (setq i 2)
          (foreach row-data sorted-data
            (setq col 0) (while (< col 5) (vla-SetCellTextHeight table-obj i col text-height) (setq col (1+ col)))
            
            ;; 공종 텍스트에서 괄호 처리
            (setq parsed-work-spec (mst-parse-work-spec (nth 1 row-data) (nth 2 row-data)))
            (setq final-work (car parsed-work-spec))
            (setq final-spec (cdr parsed-work-spec))

            (vla-SetRowHeight table-obj i (* text-height (/ 10.0 3.0)))
            (vla-settext table-obj i 0 final-work) (vla-settext table-obj i 1 final-spec) (vla-settext table-obj i 2 (nth 3 row-data))
            (setq qty-val (nth 4 row-data)) (setq qty-str (rtos qty-val 2 1)) (if (not (vl-string-search "." qty-str)) (setq qty-str (strcat qty-str ".0"))) (vla-settext table-obj i 3 qty-str)
            (vla-settext table-obj i 4 "")
            (vla-SetCellAlignment table-obj i 0 5) (vla-SetCellAlignment table-obj i 1 5) (vla-SetCellAlignment table-obj i 2 5) (vla-SetCellAlignment table-obj i 3 acAlignMiddleRight) (vla-SetCellAlignment table-obj i 4 5)
            (setq i (1+ i))
          )
          (princ " 완료.") t
        )
        (progn (princ "\n>> 표 삽입 위치가 지정되지 않아 작업을 취소합니다.") nil)
      )
    )
    (progn (princ "\n*** 오류: 집계할 데이터가 없습니다. ***") nil)
  )
)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; F. 메인 함수
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun c:MST (/ *error* old-error initial-data qty-data user-choice temp-regions processed-objects original-colors text-height scale-factor upper-offset lower-offset old-cmdecho created-group-names group-name group-dict group-obj ents-in-group ent pline texts auto-result ungrouped-texts manual-interaction-result all-initial-plines all-initial-texts text-style cleanup-flag target-layer continue-loop tables)
  (vl-load-com)
  (defun *error* (msg) (if (and msg (not (wcmatch (strcase msg) "*CANCEL*,*QUIT*"))) (princ (strcat "\n*** 오류 발생: " msg " ***"))) (if old-cmdecho (setvar "CMDECHO" old-cmdecho)) (mst-cleanup) (if old-error (setq *error* old-error)) (princ "\n작업이 중단되었습니다.")(princ))
  (setq old-error *error* *error* *error* old-cmdecho nil)
  (setq temp-regions nil processed-objects nil original-colors nil user-choice "Y" created-group-names nil all-initial-texts nil text-style nil cleanup-flag t)
  (setq *mst-temp-group-count* 0) (setq *mst-temp-numbers* nil) (setq *mst-temp-notes* nil)
  (mst-ensure-layers-on)
  (princ "\n\n*** MST (수량집계표작성) 시작 ***")
  
  (setq initial-data (mst-select-objects))
  
  (cond
    ((and (listp initial-data) (= (car initial-data) "TABLE_MODE"))
      (setq tables (cadr initial-data))
      (setq qty-data (mst-aggregate-tables tables))
      
      (if qty-data
        (progn
           (setq text-style (vla-GetTextStyle (vlax-ename->vla-object (car tables)) acDataRow))
           (setq text-height (vla-GetTextHeight (vlax-ename->vla-object (car tables)) acDataRow))
           (mst-create-table qty-data text-style "총 괄 수 량 집 계 표")
           (princ "\n\n*** 총괄수량집계표 작성 완료! ***")
        )
        (princ "\n*** 오류: 선택한 테이블에서 유효한 데이터를 찾을 수 없습니다. ***")
      )
      (setq cleanup-flag nil)
    )

    ((listp initial-data)
      (setq all-initial-plines (car initial-data))
      (setq all-initial-texts (cadr initial-data))
      (setq target-layer (caddr initial-data))
      
      (if (mst-validate-text-height all-initial-texts)
        (progn
          (setq text-style (cdr (assoc 7 (entget (car all-initial-texts)))))
          (if (mst-create-temp-regions all-initial-plines)
            (progn
              (setq auto-result (mst-find-and-color-groups (list all-initial-plines all-initial-texts) target-layer))
              (setq created-group-names (car auto-result) ungrouped-texts (cadr auto-result) processed-objects (caddr auto-result))
              
              (setq continue-loop t next-default nil)
              (while continue-loop
                (setq manual-interaction-result (mst-manual-interaction processed-objects created-group-names ungrouped-texts target-layer next-default))
                (setq next-default nil)
                (if manual-interaction-result
                  (progn
                    (cond
                      ((= (car manual-interaction-result) "FINISH")
                        (setq continue-loop nil))
                      ((= (car manual-interaction-result) "REVISE_DEFAULT")
                        (setq next-default "Revise")
                        (setq processed-objects (cadr manual-interaction-result))
                        (setq created-group-names (caddr manual-interaction-result))
                        (setq ungrouped-texts (nth 3 manual-interaction-result)))
                      (t
                        (setq processed-objects (cadr manual-interaction-result))
                        (setq created-group-names (caddr manual-interaction-result))
                        (setq ungrouped-texts (nth 3 manual-interaction-result)))
                    )
                  )
                  (setq continue-loop nil)
                )
              )

              (initget "Y N")
              (setq user-choice (getkword "\n\n>> 그룹 조정이 완료되었습니다. 계속 진행하여 표를 생성하시겠습니까? [Yes/No] <Y>: "))
              (if (or (not user-choice) (= user-choice "Y"))
                (if processed-objects
                  (if (setq qty-data (mst-analyze-groups))
                    (if (mst-create-table qty-data text-style "수 량 집 계 표") 
                      (princ "\n\n*** 수량집계표 작성 완료! ***") 
                      (setq cleanup-flag nil)
                    )
                    (setq cleanup-flag nil)
                  )
                  (princ "\n>> 경고: 유효한 그룹이 없습니다. 표를 생성할 수 없습니다.")
                )
                (progn (setq cleanup-flag nil) (princ "\n>> 작업을 중단합니다. 임시 객체는 그대로 유지됩니다."))
              )
            )
          )
        )
      )
    )
  )
  
  (if cleanup-flag (mst-cleanup))
  (setq *error* old-error)
  (princ)
)

(princ "\n>> MST.lsp 로드 완료. 명령어: MST")
(princ)

;;; ========================================================================
;;; END OF MST (통합 버전)
;;; ========================================================================
