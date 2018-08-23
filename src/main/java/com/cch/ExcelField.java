package com.cch;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD,ElementType.TYPE,ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {
	
    String value() default "";

    String title();
    
    int type() default 0;

    int align() default 0;

    int sort() default 0;
    
    int width() default -1;

    String dictType() default "";

    Class<?> fieldType() default Class.class;

    int[] groups() default {};
}
