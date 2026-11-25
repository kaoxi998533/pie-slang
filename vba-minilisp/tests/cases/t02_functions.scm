; functions, recursion, closures, higher-order functions
(define (square x) (* x x))
(display (square 12)) (newline)

(define (fact n)
  (if (= n 0)
      1
      (* n (fact (- n 1)))))
(display (fact 10)) (newline)

(define (fib n)
  (if (< n 2) n (+ (fib (- n 1)) (fib (- n 2)))))
(display (fib 15)) (newline)

; a closure remembers the environment it was created in
(define (make-adder n) (lambda (x) (+ x n)))
(define add5 (make-adder 5))
(display (add5 10)) (newline)

; set! changes a variable captured by the closure
(define (make-counter)
  (define count 0)
  (lambda ()
    (set! count (+ count 1))
    count))
(define c1 (make-counter))
(define c2 (make-counter))
(c1) (c1)
(display (list (c1) (c2))) (newline)

; functions are values
(define (twice f x) (f (f x)))
(display (twice square 3)) (newline)
(display ((lambda (a b) (- a b)) 10 3)) (newline)
(display (apply + (list 1 2 3 4))) (newline)

; a parameter that is a single symbol collects all the arguments
(define all-args (lambda args args))
(display (all-args 1 2 3)) (newline)
(display (procedure? square)) (newline)
(display (procedure? car)) (newline)
(display square) (newline)
(display car) (newline)
