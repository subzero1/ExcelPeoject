package com.csg.supervise;

import java.nio.charset.Charset;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {
//    public static String getEncoding(String str) {
//        byte[] bytes = str.getBytes();
//        Charset charset = Charset.defaultCharset();
//        return charset.displayName();
//    }
//    public static void main(String[] args) {
////        String regex="[/u][0-9a-fA-F]{4}";
//        String regex="[\\\\u]+";
//        System.out.println("regex:"+regex);
//        String input="\u16f7\uaaaa";
//        System.out.println("input:"+input);
//        boolean isMatch=input.matches(regex);
//        System.out.println(isMatch);
//    }
public static void main(String[] args) {
    String str1 = "\\u0041";
    String str2 = "\\u03B1";
    String str4 = "\\u03B1\\u03B1\\u03ff\\u06FF";
    String str3 = "Hello";

    System.out.println(str1 + " matches? " + matchesUnicode(str1));
    System.out.println(str2 + " matches? " + matchesUnicode(str2));
    System.out.println(str3 + " matches? " + matchesUnicode(str3));
    System.out.println(str4 + " matches? " + matchesUnicode(str4));
}

    public static boolean matchesUnicode(String str) {
        Pattern pattern = Pattern.compile("(\\\\u[0-9A-Fa-f]{4})+");
        Matcher matcher = pattern.matcher(str);
        return matcher.matches();
    }

}
