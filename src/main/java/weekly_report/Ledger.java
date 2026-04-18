//package lesson;
// updated on 17.04.26 @ 20.30 hrs
package weekly_report;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
public class Ledger {
	Calendar cal = Calendar.getInstance();
	SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
	public Ledger(int code, Date date, double amount) {
		// super();
		this.code = code;
		this.trndt = date;
		this.amount = amount;
	}
	public int code;
	public Date trndt;
	public double amount;
	public int getCode() {
		return code;
	}
	public void setCode(int code) {
		this.code = code;
	}
	public Date getTrndt() {
		return trndt;
	}
	public void setTrndt(Date trndt) {
		this.trndt = trndt;
	}
	public double getAmount() {
		return amount;
	}
	public void setAmount(double amount) {
		this.amount = amount;
	}
	public String age() {
		Date d = this.trndt;
		LocalDate d1 = d.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
		int a = Integer.parseInt(
				d1.getYear() + ("00" + d1.getMonthValue()).substring(("00" + d1.getMonthValue()).length() - 2));
		if (a <= 202509)
			a = 202509;
		return "" + this.code + "~" + a;
	}
	public Integer week() {
		Date d = this.trndt;
		LocalDate d1 = d.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
		String wk = "";
		switch (d1.getDayOfMonth()) {
		case 1:
		case 2:
		case 3:
		case 4:
		case 5:
		case 6:
		case 7:
			wk = "1";
			break;
		case 8:
		case 9:
		case 10:
		case 11:
		case 12:
		case 13:
		case 14:
			wk = "2";
			break;
		case 15:
		case 16:
		case 17:
		case 18:
		case 19:
		case 20:
		case 21:
			wk = "3";
			break;
		default:
			wk = "4";
			break;
		}
		return Integer.parseInt(
				d1.getYear() + ("00" + d1.getMonthValue()).substring(("00" + d1.getMonthValue()).length() - 2) + wk);
	}
}
