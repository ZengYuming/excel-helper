package com.tony.demo.excel_helper.exception;

/**
 * 正则规则校验失败异常
 */
public class ExcelRegexpValidFailedException extends Exception {
    private static final long serialVersionUID = 1L;
    private String message = null;

    public ExcelRegexpValidFailedException(Exception e) {
        initCause(e);
    }

    public ExcelRegexpValidFailedException(String message) {
        this.message = message;
    }

    public ExcelRegexpValidFailedException(String message, Exception e) {
        initCause(e);
        this.message = message;
    }

    @Override
    public String getMessage() {
        if (message == null) {
            return super.getMessage();
        }
        return message;
    }
}
