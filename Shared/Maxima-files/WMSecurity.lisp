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

;; load( ): only files inside WordMathEng's folder (where this file is) or Maxima's own installation, found
;; with Maxima's default search lists (a document can't redirect file_search_maxima/_lisp). Network (UNC)
;; paths and anything else are refused. Package names such as load(draw) or load("mnewton") still work.
(defvar *wm-sec-roots* nil)
(defvar *wm-sec-search* nil)
(defun wm-sec-norm (p)
  (string-downcase (substitute #\/ #\ (namestring p))))
(defun wm-sec-allowed-file-p (path)
  (let ((tn (ignore-errors (probe-file path))))
    (and tn
         (let ((n (wm-sec-norm tn)))
           (some (lambda (r) (and (>= (length n) (length r)) (string= r (subseq n 0 (length r)))))
                 *wm-sec-roots*)))))
(unless (get '$load 'wm-sec-orig)
  (setf *wm-sec-roots*
        (remove nil (list (and *load-truename* (wm-sec-norm (make-pathname :name nil :type nil :version nil
                                                                           :defaults *load-truename*)))
                          (and (boundp '*maxima-prefix*) *maxima-prefix*
                               (ignore-errors (wm-sec-norm (truename (concatenate 'string *maxima-prefix* "/"))))))))
  (setf *wm-sec-search* (list (copy-tree $file_search_maxima) (copy-tree $file_search_lisp)))
  (setf (get '$load 'wm-sec-orig) (symbol-function '$load))
  (setf (symbol-function '$load)
        (lambda (&rest args)
          (let* ((f (car args))
                 (fs (if (or (stringp f) (symbolp f)) (if (stringp f) f (string-downcase (subseq (symbol-name f) 1))) "")))
            (when (and (>= (length fs) 2) (member (char fs 0) '(#\/ #\\)) (member (char fs 1) '(#\/ #\\)))
              (wm-sec-refuse "load from a network path"))
            (let (($file_search_maxima (first *wm-sec-search*))
                  ($file_search_lisp (second *wm-sec-search*)))
              (let ((found (ignore-errors ($file_search f (cons '(mlist) (append (cdr $file_search_maxima) (cdr $file_search_lisp)))))))
                (unless (and (stringp found) (wm-sec-allowed-file-p found))
                  (wm-sec-refuse (format nil "load of ~a" fs))))
              (prog1 (apply (get '$load 'wm-sec-orig) args)
                ;; a loaded package may define a blocked function again
                (wm-sec-block-all)))))))

;; eval_string( ) and parse_string( ): refuse text holding ? (Lisp names such as ?run-program) or :lisp, so a
;; string built at run time can't reach Lisp. WMUnits.mac uses parse_string( ) on numbers only.
(dolist (s '($eval_string $parse_string))
  (unless (get s 'wm-sec-orig)
    (setf (get s 'wm-sec-orig) (symbol-function s))
    (let ((sym s) (wm-nm (string-downcase (subseq (symbol-name s) 1))))
      (setf (symbol-function s)
            (lambda (&rest args)
              (when (and (stringp (car args))
                         (or (find #\? (car args)) (search ":lisp" (string-downcase (car args)))))
                (wm-sec-refuse (format nil "~a with ? or :lisp" wm-nm)))
              (apply (get sym 'wm-sec-orig) args))))))
