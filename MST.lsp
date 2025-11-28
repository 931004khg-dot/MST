;;; ========================================================================
;;; MST (수량집계표작성) - 통합 로더
;;; 이 파일은 MST 프로그램의 모든 부분을 순서대로 로드합니다.
;;; ========================================================================

(defun load-mst-parts (/ part-files file-path load-success all-loaded base-path)
  (setq all-loaded t)
  (setq base-path (getvar "DWGPREFIX"))
  
  (princ "\n========================================")
  (princ "\n  MST (수량집계표작성) 로딩 중...")
  (princ "\n========================================")
  
  (setq part-files '("MST_part1.lsp" "MST_part2.lsp" "MST_part3.lsp" "MST_part4.lsp"))
  
  (foreach file part-files
    (setq file-path (strcat base-path file))
    (princ (strcat "\n로딩: " file " ... "))
    
    (if (findfile file-path)
      (progn
        (setq load-success (load file-path))
        (if load-success
          (princ "완료!")
          (progn
            (princ "실패!")
            (setq all-loaded nil)
          )
        )
      )
      (progn
        (princ "파일을 찾을 수 없습니다!")
        (setq all-loaded nil)
      )
    )
  )
  
  (if all-loaded
    (progn
      (princ "\n========================================")
      (princ "\n  MST 로딩 완료!")
      (princ "\n  명령어: MST")
      (princ "\n========================================")
      (princ)
    )
    (progn
      (princ "\n========================================")
      (princ "\n  오류: MST 로딩 실패!")
      (princ "\n  모든 part 파일이 도면과 같은 폴더에 있는지 확인하세요.")
      (princ "\n========================================")
      (princ)
    )
  )
  all-loaded
)

;; MST 부분 파일들을 로드
(load-mst-parts)

;;; ========================================================================
;;; END OF MST LOADER
;;; ========================================================================
