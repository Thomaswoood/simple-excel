package com.thomas.alib.excel.exception;

/**
 * excel解析异常
 */
public class AnalysisException extends RuntimeException {

    private String msg;

    public AnalysisException() {
        this.msg = "";
    }

    public AnalysisException(String message) {
        this.msg = message;
    }

    @Override
    public String getLocalizedMessage() {
        return msg;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }
}
