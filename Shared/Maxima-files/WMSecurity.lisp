;;; WMSecurity.lisp -- word_math security (research 08_security_audit.md)
;;; Second line of defence behind the VBA check in CMaxima.RunMaxima (WMSec_Check): Maxima functions that
;;; run operating-system commands, enter Lisp or write files are replaced, at Lisp level, by functions that
;;; refuse. Lisp-level replacements survive kill(all), remfunction and redefinition with := .
;;; Loaded by WMUnits.mac once per Maxima process. WordMathEng itself never calls these functions.
;;; Also restricted: system( ) (only gnuplot, with a checked plot script), load( ) (WordMathEng's and Maxima's
;;; own folders only), eval_string/parse_string (no ? or :lisp).
;;; Not covered here (the parser handles them before any function runs, so only the VBA check refuses them):
;;; ?name Lisp access and :lisp typed directly.

(in-package :maxima)

(defun wm-sec-refuse (name)
  (merror "WordMathEng security: ~a is not allowed." name))

(defvar *wm-sec-blocked*
  '($to_lisp $eval_string_lisp $run_testsuite $demo $batch
    $writefile $appendfile $save $stringout $with_stdout
    $opena $openw $opena_binary $openw_binary $write_data $write_binary_data
    $delete_file $rename_file $copy_file $mkdir $rmdir $chdir
    $compile_file $translate_file $compfile))

(defun wm-sec-impl (sym)
  ;; defmfun $foo also defines FOO-IMPL; compiled Lisp code (draw, plot) calls FOO-IMPL directly, so every
  ;; replacement goes on both.
  (find-symbol (concatenate 'string (subseq (symbol-name sym) 1) "-IMPL") :maxima))

(defun wm-sec-install (sym fn)
  (setf (symbol-function sym) fn)
  (let ((i (wm-sec-impl sym)))
    (when (and i (fboundp i)) (setf (symbol-function i) fn))))

(defun wm-sec-original (sym)
  ;; the function that does the work: FOO-IMPL when it exists (the $foo entry point calls it)
  (let ((i (wm-sec-impl sym)))
    (if (and i (fboundp i)) (symbol-function i) (symbol-function sym))))

(defun wm-sec-block (sym)
  ;; The function cell covers calls from Lisp, apply( ) and map( ); the mfexpr* property is looked up first by
  ;; Maxima's evaluator, so it also survives a package that redefines the function when it is autoloaded.
  (let ((refuser (or (get sym 'wm-sec-refuser)
                     (setf (get sym 'wm-sec-refuser)
                           (let ((wm-nm (string-downcase (subseq (symbol-name sym) 1))))
                             (lambda (&rest args) (declare (ignore args)) (wm-sec-refuse wm-nm)))))))
    (wm-sec-install sym refuser)
    (setf (get sym 'mfexpr*) refuser)))

(defun wm-sec-block-all ()
  (dolist (s *wm-sec-blocked*) (wm-sec-block s)))

(wm-sec-block-all)

;; Plot scripts: draw and plot2d write a gnuplot script and run gnuplot on it. Text from a document (titles,
;; labels, preambles, file names) ends up in that script, and gnuplot can run programs (system, `...`, !,
;; '< cmd', '| cmd') and read or write files. The script is checked before gnuplot runs.
(defun wm-sec-norm-path (s)
  (string-downcase (substitute #\/ #\\ s)))
(defun wm-sec-out-roots ()
  (remove nil (list (let ((tp (ignore-errors (maxima-getenv "TEMP")))) (and tp (wm-sec-norm-path tp)))
                    (and (stringp $maxima_tempdir) (wm-sec-norm-path $maxima_tempdir)))))
(defun wm-sec-out-path-ok (cmd)
  ;; cmd: "set output '<path>'" (any abbreviation); the path must be in %TEMP% or maxima_tempdir, or a plain
  ;; relative file name
  (let* ((q1 (position-if (lambda (c) (member c '(#\' #\"))) cmd))
         (q2 (and q1 (position (char cmd q1) cmd :start (1+ q1))))
         (p (and q1 q2 (wm-sec-norm-path (subseq cmd (1+ q1) q2)))))
    (and p (not (search ".." p))
         (or (and (not (find #\/ p)) (not (find #\: p)))
             (some (lambda (r) (and (> (length p) (length r)) (string= r (subseq p 0 (length r)))))
                   (wm-sec-out-roots))))))
(defun wm-sec-prefix-of (w words)
  (and (> (length w) 0) (some (lambda (d) (and (<= (length w) (length d)) (string= w (subseq d 0 (length w))))) words)))
(defun wm-sec-gnuplot-text-ok (txt)
  (let ((low (string-downcase txt)))
    (when (or (find #\` low) (find #\! low)
              (some (lambda (b) (search b low)) '("system" "shell" "popen" "eval" "import" "'<" "\"<" "'|" "\"|")))
      (return-from wm-sec-gnuplot-text-ok nil))
    (dolist (cmd (wm-sec-split low) t)
      (let* ((c (string-trim '(#\Space #\Tab) cmd))
             (words (wm-sec-words c))
             (w1 (or (first words) ""))
             (w2 (or (second words) "")))
        (cond ((member w1 '("set" "unset" "plot" "splot" "replot") :test #'string=)
               (when (and (wm-sec-prefix-of w1 '("set")) (wm-sec-prefix-of w2 '("output")))
                 (unless (wm-sec-out-path-ok c) (return nil)))
               (when (and (string= w1 "set") (>= (length w2) 2)
                          (wm-sec-prefix-of w2 '("print" "table" "loadpath" "fontpath" "psdir" "fit" "history" "logfile")))
                 (return nil)))
              ((wm-sec-prefix-of w1 '("load" "call" "cd" "save" "fit" "history" "update" "set" "print"))
               ;; abbreviations (se o '...') and file commands
               (when (wm-sec-prefix-of w1 '("set"))
                 (when (wm-sec-prefix-of w2 '("output")) (unless (wm-sec-out-path-ok c) (return nil)))
                 (when (>= (length w2) 2)
                   (when (wm-sec-prefix-of w2 '("print" "table" "loadpath" "fontpath" "psdir" "fit" "history" "logfile"))
                     (return nil))))
               (unless (or (wm-sec-prefix-of w1 '("set")) (wm-sec-prefix-of w1 '("print")))
                 (return nil))))))))
(defun wm-sec-split (s)
  ;; commands: separated by newlines and ; (every ; counts, also inside quotes: conservative)
  (let ((out nil) (start 0))
    (dotimes (i (length s))
      (when (member (char s i) '(#\Newline #\Return #\; #\{ #\}))
        (push (subseq s start i) out) (setf start (1+ i))))
    (push (subseq s start) out)
    (nreverse out)))
(defun wm-sec-words (s)
  (let ((out nil) (cur nil))
    (loop for ch across s do
      (if (member ch '(#\Space #\Tab #\( #\{))
          (when cur (push (coerce (nreverse cur) 'string) out) (setf cur nil))
          (push ch cur)))
    (when cur (push (coerce (nreverse cur) 'string) out))
    (nreverse out)))
(defun wm-sec-read-file (p)
  (ignore-errors
    (with-open-file (s p :if-does-not-exist nil :external-format :latin-1)
      (when s (let ((str (make-string (file-length s)))) (subseq str 0 (read-sequence str s)))))))
(defun wm-sec-gnuplot-arg-ok (a)
  (and (stringp a)
       (not (find-if (lambda (c) (find c "&|<>^%!`")) a))
       (not (find #\Newline a)) (not (find #\Return a))
       (or (string= a "-persist")
           (let ((txt (wm-sec-read-file (string-trim "\" " a))))
             (and txt (wm-sec-gnuplot-text-ok txt))))))

;; system( ) is needed by Maxima's own plotting (graphs, drawdf), which calls it from Lisp with gnuplot_command
;; and the plot script. Only that command, as it was when this file was loaded, is let through, with arguments
;; free of cmd.exe metacharacters and a plot script that passes wm-sec-gnuplot-text-ok (above).
;; Typed in a document, system( ) is always refused (mfexpr* below).
(unless (get '$system 'wm-sec-orig)
  (setf (get '$system 'wm-sec-orig) (wm-sec-original '$system))
  (setf (get '$system 'wm-sec-gnuplot)
        (if (and (boundp '$gnuplot_command) (stringp $gnuplot_command)) (copy-seq $gnuplot_command) nil))
  (wm-sec-install '$system
        (lambda (&rest args)
          (let ((ok (get '$system 'wm-sec-gnuplot)))
            (if (and ok args (stringp (car args)) (string= (car args) ok)
                     (every #'wm-sec-gnuplot-arg-ok (cdr args)))
                (apply (get '$system 'wm-sec-orig) args)
                (wm-sec-refuse "system")))))
  (setf (get '$system 'mfexpr*) (lambda (&rest args) (declare (ignore args)) (wm-sec-refuse "system"))))

;; load( ): only files inside WordMathEng's folder (where this file is) or Maxima's own installation, found
;; with Maxima's default search lists (a document can't redirect file_search_maxima/_lisp). Network (UNC)
;; paths and anything else are refused. Package names such as load(draw) or load("mnewton") still work.
(defvar *wm-sec-roots* nil)
(defvar *wm-sec-search* nil)
(defun wm-sec-norm (p)
  (string-downcase (substitute #\/ #\\ (namestring p))))
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
  (setf (get '$load 'wm-sec-orig) (wm-sec-original '$load))
  (wm-sec-install '$load
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
    (setf (get s 'wm-sec-orig) (wm-sec-original s))
    (let ((sym s) (wm-nm (string-downcase (subseq (symbol-name s) 1))))
      (wm-sec-install s
            (lambda (&rest args)
              (when (and (stringp (car args))
                         (or (find #\? (car args)) (search ":lisp" (string-downcase (car args)))))
                (wm-sec-refuse (format nil "~a with ? or :lisp" wm-nm)))
              (apply (get sym 'wm-sec-orig) args))))))
