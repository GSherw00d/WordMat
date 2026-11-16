;;; WMSecurity.lisp -- word_math security (research 08_security_audit.md)
;;; Second line of defence behind the VBA check in CMaxima.RunMaxima (WMSec_Check): Maxima functions that
;;; run operating-system commands, enter Lisp or write files are replaced, at Lisp level, by functions that
;;; refuse. Lisp-level replacements survive kill(all), remfunction and redefinition with := .
;;; Loaded by WMUnits.mac once per Maxima process. WordMathEng itself never calls these functions.
;;; Not covered here (WordMathEng needs them, so only the VBA check refuses them from documents):
;;; load/batchload (except UNC network paths, refused below), ?name Lisp access, :lisp.

(in-package :maxima)

(defun wm-sec-refuse (name)
  (merror "WordMathEng security: ~a is not allowed." name))

(defvar *wm-sec-blocked*
  '($to_lisp $eval_string_lisp $run_testsuite $demo $batch
    $writefile $appendfile $save $stringout $with_stdout
    $opena $openw $opena_binary $openw_binary $write_data $write_binary_data
    $delete_file $rename_file $copy_file $mkdir $rmdir $chdir
    $compile_file $translate_file $compfile))

(defun wm-sec-block (sym)
  ;; The function cell covers calls from Lisp, apply( ) and map( ); the mfexpr* property is looked up first by
  ;; Maxima's evaluator, so it also survives a package that redefines the function when it is autoloaded.
  (let ((refuser (or (get sym 'wm-sec-refuser)
                     (setf (get sym 'wm-sec-refuser)
                           (let ((wm-nm (string-downcase (subseq (symbol-name sym) 1))))
                             (lambda (&rest args) (declare (ignore args)) (wm-sec-refuse wm-nm)))))))
    (unless (and (fboundp sym) (eq (symbol-function sym) refuser))
      (setf (symbol-function sym) refuser))
    (setf (get sym 'mfexpr*) refuser)))

(defun wm-sec-block-all ()
  (dolist (s *wm-sec-blocked*) (wm-sec-block s)))

(wm-sec-block-all)

;; system( ) is needed by Maxima's own plotting (drawdf), which calls it from Lisp with gnuplot_command and the
;; plot file. Only that command, as it was when this file was loaded, is let through, and only with arguments
;; free of cmd.exe metacharacters (gnuplot_file_args is a Maxima variable a document could change).
;; Typed in a document, system( ) is always refused (mfexpr* below).
(unless (get '$system 'wm-sec-orig)
  (setf (get '$system 'wm-sec-orig) (symbol-function '$system))
  (setf (get '$system 'wm-sec-gnuplot)
        (if (and (boundp '$gnuplot_command) (stringp $gnuplot_command)) (copy-seq $gnuplot_command) nil))
  (setf (symbol-function '$system)
        (lambda (&rest args)
          (let ((ok (get '$system 'wm-sec-gnuplot)))
            (if (and ok args (stringp (car args)) (string= (car args) ok)
                     (every (lambda (a) (and (stringp a)
                                             (not (find-if (lambda (c) (find c "&|<>^%!`")) a))
                                             (not (find #\Newline a)) (not (find #\Return a))))
                            (cdr args)))
                (apply (get '$system 'wm-sec-orig) args)
                (wm-sec-refuse "system")))))
  (setf (get '$system 'mfexpr*) (lambda (&rest args) (declare (ignore args)) (wm-sec-refuse "system"))))

;; load( ): refuse network (UNC) paths such as //server/share/file.mac
(unless (get '$load 'wm-sec-orig)
  (setf (get '$load 'wm-sec-orig) (symbol-function '$load))
  (setf (symbol-function '$load)
        (lambda (&rest args)
          (let ((f (car args)))
            (when (and (stringp f) (>= (length f) 2)
                       (member (char f 0) '(#\/ #\\)) (member (char f 1) '(#\/ #\\)))
              (wm-sec-refuse "load from a network path"))
            (prog1 (apply (get '$load 'wm-sec-orig) args)
              ;; a loaded package may define a blocked function again
              (wm-sec-block-all))))))
