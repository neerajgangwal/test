/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.exceljava.jinx.examples;

import com.exceljava.jinx.ExcelFunction;


/**
 *
 * @author neeraj
 */
public class Test {
     /***
     * Multiply two numbers and return the result.
     * @param x
     * @param y
     * @return 
     */
    @ExcelFunction
    public static double multiply(double x, double y) {
        return x * y;
    }
}