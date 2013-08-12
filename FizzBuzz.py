#!/usr/bin/env python

def fizzbuzz(n):
    i = 1
    while i <= n:
        r = i
        if i % 3 == 0:
            r = 'Fizz'
        if i % 5 == 0:
            r = 'Buzz'
        if i % 3 == 0 and i % 5 == 0:
            r = 'FizzBuzz'
        print r
        i = i + 1      
                    
print fizzbuzz(100)