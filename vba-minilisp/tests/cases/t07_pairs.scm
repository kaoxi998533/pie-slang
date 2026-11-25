; real pairs: dotted pairs, sharing, mutation, identity
(define p (cons 1 2))
(display p) (newline)
(display (car p)) (display " ") (display (cdr p)) (newline)
(display (pair? p)) (display (pair? '())) (display (list? p)) (newline)
(display '(1 . (2 . (3 . ())))) (newline)
(display '(a b . c)) (newline)
(display (cons 1 (cons 2 3))) (newline)
(display (cons '(1 2) '(3 4))) (newline)

; an association list (a list of pairs)
(define phone '((ann . 1234) (bob . 5678)))
(define (lookup key alist)
  (cond ((null? alist) #f)
        ((eq? key (car (car alist))) (cdr (car alist)))
        (else (lookup key (cdr alist)))))
(display (lookup 'bob phone)) (display " ") (display (lookup 'cy phone)) (newline)

; cons only adds one pair: the tail is shared, not copied
(define tail '(2 3))
(define a (cons 1 tail))
(define b (cons 0 tail))
(set-car! tail 99)
(display a) (display " ") (display b) (newline)
(display (eq? (cdr a) (cdr b))) (newline)

; eq? is identity, equal? compares contents
(define x (list 1 2))
(define y (list 1 2))
(display (list (eq? x y) (equal? x y) (eq? x x) (eq? '() '()))) (newline)

; set-cdr! builds structure in place: a queue with a tail pointer
(define (make-queue) (cons '() '()))
(define (enqueue! q item)
  (define cell (list item))
  (if (null? (car q))
      (begin (set-car! q cell) (set-cdr! q cell))
      (begin (set-cdr! (cdr q) cell) (set-cdr! q cell))))
(define q (make-queue))
(enqueue! q 'a) (enqueue! q 'b) (enqueue! q 'c)
(display (car q)) (newline)

; rest parameters use dotted parameter lists
(define (f a . rest) (list a rest))
(display (f 1)) (display (f 1 2 3)) (newline)
(define (count . items) (length items))
(display (count 'a 'b 'c)) (newline)
(display ((lambda (a b . c) (list a b c)) 1 2 3 4)) (newline)

; append copies all but the last list
(display (append '(1) '(2) 3)) (newline)
(display (append)) (newline)
(display (apply + '(1 2 3))) (newline)
