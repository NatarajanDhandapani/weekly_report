//package lesson;
// updated on 17.04.26 @ 20.30 hrs
package weekly_report;

public class Master {
	public Master(int code, String custname, String zm, String csm, int r, double plan, double actual) {
		super();
		this.code = code;
		this.custname = custname;
		this.zm = zm;
		this.csm = csm;
		this.r = r;
		this.plan = plan;
		this.actual = actual;
	}
	public int getCode() {
		return code;
	}
	public void setCode(int code) {
		this.code = code;
	}
	public String getCustname() {
		return custname;
	}
	public void setCustname(String custname) {
		this.custname = custname;
	}
	public String getZm() {
		return zm;
	}
	public int getr() {
		return r;
	}
	public double getplan() {
		return plan;
	}
	public double getactual() {
		return actual;
	}
	public void setZm(String zm) {
		this.zm = zm;
	}
	public String getCsm() {
		return csm;
	}
	public void setCsm(String csm) {
		this.csm = csm;
	}
	public void setr(int r) {
		this.r = r;
	}
	public void setplan(double plan) {
		this.plan = plan;
	}
	public void setactual(double actual) {
		this.actual = actual;
	}
	public String getAll() {
		return "" + this.zm + "~" + this.csm + "~" + this.code + "~" + this.custname + "~" + this.r + "~" + this.plan
				+ "~" + this.actual;
	}
	public int code;
	public String custname;
	public String zm;
	public String csm;
	public int r;
	public double plan;
	public double actual;
}
