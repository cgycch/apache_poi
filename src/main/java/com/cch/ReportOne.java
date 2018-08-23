package com.cch;

public class ReportOne {
	@ExcelField(title="name",sort=1)
	private String name;
	@ExcelField(title="addr",sort=2)
	private String toAddr;
	@ExcelField(title="age",sort=3)
	private String age;
	
	
	public ReportOne() {
		super();
	}
	public ReportOne(String name, String toAddr, String age) {
		super();
		this.name = name;
		this.toAddr = toAddr;
		this.age = age;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getToAddr() {
		return toAddr;
	}
	public void setToAddr(String toAddr) {
		this.toAddr = toAddr;
	}
	public String getAge() {
		return age;
	}
	public void setAge(String age) {
		this.age = age;
	}
	
	

}
