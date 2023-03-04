package com.thomas.alib.excel.importer;

/**
 * excel导入安全模式返回类型
 *
 * @param <T> 读取的数据类型的泛型
 */
public class SafetyResult<T> {
    int rowIndex;
    boolean readSuccess;
    StringBuffer msg;
    T data;

    public SafetyResult() {
        readSuccess = true;//默认读取成功设为真，当有读取失败时通过setter设置为假
        msg = new StringBuffer();
    }

    public SafetyResult(int rowIndex) {
        this();
        this.rowIndex = rowIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public boolean isReadSuccess() {
        return readSuccess;
    }

    public void setReadSuccess(boolean readSuccess) {
        this.readSuccess = readSuccess;
    }

    public String getMsg() {
        return msg.toString();
    }

    public void setMsg(StringBuffer msg) {
        this.msg = msg;
    }

    public SafetyResult<T> appendMsg(String msg) {
        this.msg.append(msg);
        return this;
    }

    public SafetyResult<T> appendMsg(Object obj) {
        this.msg.append(obj);
        return this;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }
}
