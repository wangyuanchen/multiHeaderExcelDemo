package com.example.demo;


import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {
    public static void main(String[] args) {
        // 匹配汉语正则表达式
        String s = "[\\u4e00-\\u9fa5]";
        Pattern r = Pattern.compile("[\\u4e00-\\u9fa5]");
        String ss = "于123";
        Matcher m = r.matcher(ss);
        if (!m.matches()) {
            System.out.println("报错"
            );
        }
        
    }
}
