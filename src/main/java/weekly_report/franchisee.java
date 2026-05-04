package weekly_report;

//updated on 03.05.26 @ 10.30 am
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.concurrent.TimeUnit;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.stream.Collectors;
import java.time.LocalDateTime;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BorderExtent;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class franchisee {
	@SuppressWarnings("resource")
	public static void main(String[] args) throws Exception {
		System.out.println("updated on " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd-MM-yy HH:mm:ss"))
				+ " " + System.getProperty("user.dir"));
		Date start = new Date();
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Summary");
		CellStyle Title = null;
		Font font = wb.createFont();
		font.setFontHeightInPoints((short) 13);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		Title = wb.createCellStyle();
		Title.setFont(font);
		CellStyle header = null;
		font = wb.createFont();
		font.setFontHeightInPoints((short) 12);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setItalic(false);
		header = wb.createCellStyle();
		header.setFont(font);
		header.setAlignment(HorizontalAlignment.LEFT);
		CellStyle header2 = null;
		font = wb.createFont();
		font.setFontHeightInPoints((short) 12);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		header2 = wb.createCellStyle();
		header2.setFont(font);
		header2.setAlignment(HorizontalAlignment.LEFT);
		CellStyle body = null;
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(false);
		font.setItalic(false);
		body = wb.createCellStyle();
		body.setFont(font);
		CellStyle amount = wb.createCellStyle();
		DataFormat format = wb.createDataFormat();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(false);
		font.setItalic(false);
		amount = wb.createCellStyle();
		amount.setDataFormat(format.getFormat("#,##,##0"));
		amount.setFont(font);
		CellStyle amount1 = wb.createCellStyle();
		format = wb.createDataFormat();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(false);
		font.setItalic(false);
		amount1 = wb.createCellStyle();
		amount1.setDataFormat(format.getFormat("#,##,##0.00"));
		amount1.setFont(font);
		CellStyle amount2 = wb.createCellStyle();
		format = wb.createDataFormat();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		amount2 = wb.createCellStyle();
		amount2.setDataFormat(format.getFormat("#,##,##0.00"));
		amount2.setFont(font);
		CellStyle total = wb.createCellStyle();
		format = wb.createDataFormat();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setItalic(false);
		total = wb.createCellStyle();
		total.setDataFormat(format.getFormat("#,##,##0"));
		total.setFont(font);
		CellStyle rightAligned = wb.createCellStyle();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 12);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setItalic(false);
		rightAligned.setAlignment(HorizontalAlignment.RIGHT);
		rightAligned.setFont(font);
		CellStyle rightAligned2 = wb.createCellStyle();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 12);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		rightAligned2.setAlignment(HorizontalAlignment.RIGHT);
		rightAligned2.setFont(font);
		CellStyle rightAlignedbody = wb.createCellStyle();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		rightAlignedbody.setAlignment(HorizontalAlignment.RIGHT);
		rightAlignedbody.setFont(font);
		String[] head = { "Code", "Customer", "ZM", "CSM",  "<Sep-25",   "Oct-25", "Nov-25", "Dec-25", 
				"Jan-26","Feb-26","Mar-26","Apr-26","May-26","Jun-26", "Jul-26","Aug-26","Sep-26",   "UAC", "Due", "Not Due", "Total",
				"Plan-Apr", "Act 1-30" };
		Map<String, String> duedates = new LinkedHashMap<String, String>();
		duedates.put("Code", "");
		duedates.put("Name", "");
		
		
		duedates.put("Feb-1st", "21.02.26");
		duedates.put("Feb-2nd", "28.02.26");
		duedates.put("Feb-3rd", "07.03.26");
		duedates.put("Feb-4th", "14.03.26");
		
		duedates.put("Mar-1st", "21.03.26");
		duedates.put("Mar-2nd", "28.03.26");
		duedates.put("Mar-3rd", "07.04.26");
		duedates.put("Mar-4th", "14.04.26");
		
		duedates.put("Apr-1st", "21.04.26");
		duedates.put("Apr-2nd", "28.04.26");
		duedates.put("Apr-3rd", "07.05.26");
		duedates.put("Apr-4th", "14.05.26");
		
		
		Set<String> head1 = new LinkedHashSet<String>(duedates.keySet());
		Set<String> head2 = new LinkedHashSet<String>(duedates.values());
		String csm[] = { "Blore", "Jeevan", "Niyamathullah", "Prakash(TN)", "Prakash", "Total", "Recovery",
				"Grand Total" };
		String[] tap0, tap1;
		int out = 0, in = 0, row = 5, col = 1;
		double zero = 0.00, notdue = 0;
		double plan = 0.00, actual = 0.00;
		boolean uac = false;
		SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy");
		ZoneId defaultZoneId = ZoneId.systemDefault();
		String notduefm = "15.04.2026"; // --Old 1-8-15-22
		String trndt = "30.09.2025"; // first column - upto < mmmYY
		LocalDate dd = ft.parse(notduefm).toInstant().atZone(ZoneId.systemDefault()).toLocalDate().minusDays(1);
		Date d3 = Date.from(dd.atStartOfDay(defaultZoneId).toInstant());
		List<Ledger> led = new ArrayList<Ledger>();
		List<Ledger> unpaid = new ArrayList<Ledger>();
		List<Master> zm = new ArrayList<Master>();
		FileInputStream xlsfile = new FileInputStream("d:/tnstore/Ledger.xlsx");
		Workbook xls = WorkbookFactory.create(xlsfile);
		System.out.println("updating customer list  ...");
		Sheet sht = xls.getSheetAt(0);
		for (Row r : sht) {
	 		if (r.getRowNum() > 0) {
				{
					Master a1 = new Master((int) r.getCell(0).getNumericCellValue(), r.getCell(1).getStringCellValue(),
							r.getCell(2).getStringCellValue(), r.getCell(3).getStringCellValue(),
							(int) r.getCell(4).getNumericCellValue(), r.getCell(6).getNumericCellValue(),
							r.getCell(7).getNumericCellValue());
					zm.add(a1);
				}
			}
		}
		int error = 0;
		System.out.println("updating Ledger ...");
		sht = xls.getSheetAt(1);
		for (Row r : sht) {
			if (r.getRowNum() > 0) {
				String temp = trndt;
				switch (r.getCell(2).getStringCellValue().trim()) {
				case "RV":
				case "DR":
					Ledger bills = new Ledger((int) r.getCell(1).getNumericCellValue(),
							ft.parse(r.getCell(5).getStringCellValue()), r.getCell(8).getNumericCellValue());
					unpaid.add(bills);
				case "SA": {
					temp = ft.parse(r.getCell(5).getStringCellValue()).before(ft.parse(notduefm))
							? r.getCell(5).getStringCellValue()
							: "31.03.2099";
					uac = true;
					break;
				}
				case "DZ":
					temp = ft.parse(r.getCell(5).getStringCellValue()).after(ft.parse(trndt)) ? "31.03.3099"
							: r.getCell(5).getStringCellValue();
					break;
				default:
					temp = r.getCell(5).getStringCellValue();
					uac = true;
				}
				if (uac && ft.parse(temp).compareTo(ft.parse(trndt)) > 0 && r.getCell(8).getNumericCellValue() < 0)
					temp = "31.03.2099";
				Ledger a1 = new Ledger((int) r.getCell(1).getNumericCellValue(), ft.parse(temp),
						r.getCell(8).getNumericCellValue());
				led.add(a1);
				uac = false;
				System.out.println(error++);
			}
		}
		System.out.println("ledger uploaded completed   ...");
		Set<Date> d2 = (Set<Date>) unpaid.stream().map(u -> u.trndt).collect(Collectors.toCollection(TreeSet::new));
		Date d = (Date) (d2.toArray()[d2.size() - 1]); // old 3
		Map<String, Double> ageing = led.stream()
				.collect(Collectors.groupingBy(Ledger::age, TreeMap::new, Collectors.summingDouble(Ledger::getAmount)));
		// @SuppressWarnings("resource")
		Map<Integer, Map<Integer, Double>> nested = led.stream().filter(a -> {
			try {
				return (a.trndt.after(ft.parse("31.01.2026")) && a.trndt.before(ft.parse("30.04.2026"))
						&& a.getAmount() > 0);
			} catch (ParseException e) {
				e.printStackTrace();
			}
			return false;
		}).collect(Collectors.groupingBy(Ledger::getCode, TreeMap::new,
				Collectors.groupingBy(Ledger::week, Collectors.summingDouble(Ledger::getAmount))));
		HashMap<Integer, String> code = (HashMap<Integer, String>) zm.stream()
				.collect(Collectors.toMap(Master::getCode, Master::getCustname));
		Set<String> zmname = zm.stream().map(a -> a.getAll()).collect(Collectors.toCollection(TreeSet::new));
		row = 5;
		TreeMap<String, Integer> mh = new TreeMap<>();
		
		
		mh.put("202509", row++);
		mh.put("202510", row++);
		mh.put("202511", row++);
		mh.put("202512", row++);
		mh.put("202601", row++);
		mh.put("202602", row++);
		mh.put("202603", row++);
		mh.put("202604", row++);
		mh.put("202605", row++);
		mh.put("202606", row++);
		mh.put("202607", row++);
		mh.put("202608", row++);
		mh.put("202609", row++);
		
		
		
		mh.put("309903", 18); /* not to correct old 15 */
		mh.put("209903", 20); /* not to correct old 17 */
		row = 5;
		for (String s : zmname) {
			zero = 0.00;
			notdue = 0.00;
			tap0 = s.split("~");
			// row = Integer.parseInt(tap0[4]);
			out = Integer.parseInt(tap0[2]);
			row = Integer.parseInt(tap0[4]);
			plan = Double.parseDouble(tap0[5]);
			actual = Double.parseDouble(tap0[6]);
			Row r = sheet.createRow(row);
			for (Entry<String, Double> a : ageing.entrySet()) {
				{
					tap1 = a.getKey().split("~");
					in = Integer.parseInt(tap1[0]);
					if (out == in) {
						r.createCell(col).setCellValue(Integer.parseInt(tap0[2]));
						r.getCell(col++).setCellStyle(body);
						r.createCell(col).setCellValue(tap0[3]);
						r.getCell(col++).setCellStyle(body);
						r.createCell(col).setCellValue(tap0[0]);
						r.getCell(col++).setCellStyle(body);
						r.createCell(col).setCellValue(tap0[1]);
						r.getCell(col++).setCellStyle(body);
						zero = mh.get(tap1[1]) == 5 && a.getValue() <= 0 ? 0.001 : a.getValue();
						if (mh.get(tap1[1]) == 5) {
							notdue = a.getValue() < 0 ? a.getValue() : 0.00;
							r.createCell(mh.get(tap1[1])).setCellValue(zero);
							r.getCell(mh.get(tap1[1])).setCellStyle(amount);
						}
						if (mh.get(tap1[1]) > 5 && mh.get(tap1[1]) < 20) { // old
																			// 17
							r.createCell(mh.get(tap1[1])).setCellValue(a.getValue());
							r.getCell(mh.get(tap1[1])).setCellStyle(amount);
						}
						if (mh.get(tap1[1]) == 20) { // old 17
							r.createCell(mh.get(tap1[1])).setCellValue(a.getValue() + notdue);
							r.getCell(mh.get(tap1[1])).setCellStyle(amount);
						}
						r.createCell(22).setCellValue(plan);
						r.getCell(22).setCellStyle(amount);
						r.createCell(23).setCellValue(actual);
						r.getCell(23).setCellStyle(amount);
					}
					col = 1;
				}
			}
			row++;
		}
		// Horizontal total
		int[] blank = { 12, 40, 71, 81, 87, 104 };
		for (int r = 5; r < sheet.getLastRowNum() + 1; r++) {
			if (r == blank[0] || r == blank[1] || r == blank[2] || r == blank[3] || r == blank[4])
				r++;
			// System.out.println(r);
			sheet.getRow(r).createCell(19).setCellFormula("sum(f" + (r + 1) + ":s" + (r + 1) + ")");
			sheet.getRow(r).getCell(19).setCellStyle(amount);
			sheet.getRow(r).createCell(21).setCellFormula("sum(t" + (r + 1) + ":u" + (r + 1) + ")");
			sheet.getRow(r).getCell(21).setCellStyle(amount);
		}
		row = 2;
		col = 1;
		sheet.createRow(row).createCell(col).setCellValue("Franchisee outstanding as on " + ft.format(d));
		sheet.getRow(row).getCell(col).setCellStyle(Title);
		row = 4;
		Row r = sheet.createRow(row);
		for (col = 1; col < head.length + 1; col++) {
			r.createCell(col).setCellValue(head[col - 1]);
			r.getCell(col).setCellStyle((col < 5) ? header : rightAligned);
		}
		// vertical total
		int sl = sheet.getLastRowNum();
		Row row1 = sheet.createRow(sl + 2);
		for (col = 5; col < head.length + 1; col++) {
			row1.createCell(col).setCellFormula("sum(" + (char) (97 + col) + "6:" + (char) (97 + col) + (sl + 1) + ")");
			row1.getCell(col).setCellStyle(total);
		}
		// printing zm summary
		row = 115;
		Row row2 = sheet.createRow(row);
		for (col = 4; col < head.length + 1; col++) {
			row2.createCell(col).setCellValue(head[col - 1]);
			row2.getCell(col).setCellStyle((col < 5) ? header2 : rightAligned2);
		}
		row = 116;
		for (int r1 = 0; r1 < csm.length; r1++) {
			sheet.createRow(row).createCell(4).setCellValue(csm[r1]);
			sheet.getRow(row++).getCell(4).setCellStyle(body);
		}
		row = 116;
		for (col = 5; col < head.length + 1; col++) {
			sheet.getRow(row).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "6:" + (char) (97 + col) + 12 + ")/10^5");
			sheet.getRow(row).getCell(col).setCellStyle(amount1);
		}
		row = 117;
		for (col = 5; col < head.length + 1; col++) {
			sheet.getRow(row).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "14:" + (char) (97 + col) + 40 + ")/10^5");
			sheet.getRow(row).getCell(col).setCellStyle(amount1);
		}
		row++;
		for (col = 5; col < head.length + 1; col++) {
			sheet.getRow(row).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "42:" + (char) (97 + col) + 71 + ")/10^5");
			sheet.getRow(row).getCell(col).setCellStyle(amount1);
		}
		row++;
		for (col = 5; col < head.length + 1; col++) {
			sheet.getRow(row).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "73:" + (char) (97 + col) + 81 + ")/10^5");
			sheet.getRow(row).getCell(col).setCellStyle(amount1);
		}
		row++;
		for (col = 5; col < head.length + 1; col++) {
			sheet.getRow(row).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "83:" + (char) (97 + col) + 87 + ")/10^5");
			sheet.getRow(row).getCell(col).setCellStyle(amount1);
		}
		row++;
		for (col = 5; col < head.length + 1; col++) {
			sheet.getRow(row).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "117:" + (char) (97 + col) + 121 + ")");
			sheet.getRow(row).getCell(col).setCellStyle(amount2);
		}
		row++;
		for (col = 5; col < head.length + 1; col++) {
			sheet.getRow(row).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "89:" + (char) (97 + col) + 97 + ")/10^5");
			sheet.getRow(row).getCell(col).setCellStyle(amount1);
		}
		row++;
		for (col = 5; col < head.length + 1; col++) {
			sheet.getRow(row).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "122:" + (char) (97 + col) + 123 + ")");
			sheet.getRow(row).getCell(col).setCellStyle(amount2);
		}
		row++;
		sheet.createRow(87).createCell(2).setCellValue("Inactive / Recovery / Commercial suit (excl rental dues)");
		sheet.getRow(87).getCell(2).setCellStyle(header2);
		sheet.getRow(115).createCell(24).setCellValue("%");
		sheet.getRow(115).getCell(24).setCellStyle(rightAligned2);
		for (row = 117; row < 125; row++) {
			sheet.getRow(row - 1).createCell(24)
					.setCellFormula(" sum( " + (char) (120) + row + "/" + (char) (119) + row + ")*100");
			sheet.getRow(row - 1).getCell(24).setCellStyle(amount1);
		}
		// printing weekly details
		row = 3;
		TreeMap<Integer, Integer> wh = new TreeMap<Integer, Integer>();
		
		
		wh.put(2026021, row++);
		wh.put(2026022, row++);
		wh.put(2026023, row++);
		wh.put(2026024, row++);
		
		wh.put(2026031, row++);
		wh.put(2026032, row++);
		wh.put(2026033, row++);
		wh.put(2026034, row++);
		
		wh.put(2026041, row++);
		wh.put(2026042, row++);
		wh.put(2026043, row++);
		wh.put(2026044, row++);
		
		
		XSSFSheet sheet1 = wb.createSheet("Weekwise");
		r = sheet1.createRow(2);
		sl = 1;
		for (String s : head1) {
			r.createCell(sl).setCellValue(s);
			r.getCell(sl++).setCellStyle(rightAligned);
		}
		r = sheet1.createRow(3);
		sl = 2;
		for (String s : head2) {
			r.createCell(sl).setCellValue(s);
			r.getCell(sl++).setCellStyle(rightAlignedbody);
		}
		for (Entry<Integer, Map<Integer, Double>> a : nested.entrySet()) {
			{
				r = sheet1.createRow(row++);
				r.createCell(1).setCellValue(a.getKey());
				r.createCell(2).setCellValue(code.get(a.getKey()));
				r.getCell(1).setCellStyle(body);
				r.getCell(2).setCellStyle(body);
				for (Entry<Integer, Double> b : a.getValue().entrySet()) {
					r.createCell(wh.get(b.getKey())).setCellValue(b.getValue());
					r.getCell(wh.get(b.getKey())).setCellStyle(amount);
				}
			}
		} // vertical total
		sl = sheet1.getLastRowNum();
		row1 = sheet1.createRow(sl + 2);
		for (col = 3; col < head1.size() + 1; col++) {
			// System.out.println(col);
			row1.createCell(col).setCellFormula("sum(" + (char) (97 + col) + "5:" + (char) (97 + col) + (sl + 1) + ")");
			row1.getCell(col).setCellStyle(total);
		}
		// printing weekly from unpaid bills
		// @SuppressWarnings("resource")
		Map<Integer, Map<Integer, Double>> unpaidbills = unpaid.stream().filter(a -> {
			try {
				return (a.trndt.after(d3) && a.getAmount() > 0);
			} catch (Exception e) {
				e.printStackTrace();
			}
			return false;
		}).collect(Collectors.groupingBy(Ledger::getCode, TreeMap::new,
				Collectors.groupingBy(Ledger::week, Collectors.summingDouble(Ledger::getAmount))));
		XSSFSheet sheet2 = wb.createSheet("Not due");
		r = sheet2.createRow(2);
		row = 4;
		for (Entry<Integer, Map<Integer, Double>> a : unpaidbills.entrySet()) {
			{
				r = sheet2.createRow(row++);
				r.createCell(1).setCellValue(a.getKey());
				r.createCell(2).setCellValue(code.get(a.getKey()));
				r.getCell(1).setCellStyle(body);
				r.getCell(2).setCellStyle(body);
				for (Entry<Integer, Double> b : a.getValue().entrySet()) {
					System.out.println(b.getKey());
					r.createCell(wh.get(b.getKey())).setCellValue(b.getValue());
					r.getCell(wh.get(b.getKey())).setCellStyle(amount);
				}
			}
		}
		r = sheet2.createRow(2);
		sl = 1;
		for (String s : head1) {
			r.createCell(sl).setCellValue(s);
			r.getCell(sl++).setCellStyle(rightAligned);
		}
		r = sheet2.createRow(3);
		sl = 2;
		for (String s : head2) {
			r.createCell(sl).setCellValue(s);
			r.getCell(sl++).setCellStyle(rightAlignedbody);
		}
		// vertical total
		sl = sheet2.getLastRowNum();
		row1 = sheet2.createRow(sl + 2);
		for (col = 3; col < head1.size() + 1; col++) {
			row1.createCell(col).setCellFormula("sum(" + (char) (97 + col) + "5:" + (char) (97 + col) + (sl + 1) + ")");
			row1.getCell(col).setCellStyle(total);
		}
		// closing the files
		PropertyTemplate pt = new PropertyTemplate();
		pt.drawBorders(new CellRangeAddress(4, 100, 1, sheet.getRow(4).getLastCellNum() - 1), BorderStyle.THIN,
				IndexedColors.GREY_25_PERCENT.getIndex(), BorderExtent.ALL);
		pt.applyBorders(sheet);
		pt = new PropertyTemplate();
		pt.drawBorders(new CellRangeAddress(115, sheet.getLastRowNum(), 3, sheet.getRow(115).getLastCellNum() - 1),
				BorderStyle.THIN, IndexedColors.GREY_25_PERCENT.getIndex(), BorderExtent.ALL);
		pt.applyBorders(sheet);
		pt = new PropertyTemplate();
		pt.drawBorders(new CellRangeAddress(2, sheet1.getLastRowNum(), 1, sheet1.getRow(2).getLastCellNum() - 1),
				BorderStyle.THIN, IndexedColors.GREY_25_PERCENT.getIndex(), BorderExtent.ALL);
		pt.applyBorders(sheet1);
		pt = new PropertyTemplate();
		pt.drawBorders(new CellRangeAddress(2, sheet2.getLastRowNum(), 1, sheet2.getRow(2).getLastCellNum() - 1),
				BorderStyle.THIN, IndexedColors.GREY_25_PERCENT.getIndex(), BorderExtent.ALL);
		pt.applyBorders(sheet2);
		sheet.setDisplayGridlines(false);
		sheet1.setDisplayGridlines(false);
		sheet2.setDisplayGridlines(false);
		sheet.setColumnWidth(0, 500); 
		sheet.setColumnWidth(1, 1800);
		sheet.setColumnWidth(2, 8800);
		sheet.setColumnWidth(3, 3200);
		sheet.setColumnWidth(4, 3300);
		sheet.setColumnWidth(5, 2650);
		sheet.setColumnHidden(3, true); // zm name column
		for (int cc = 6; cc < head.length + 1; cc++)
			sheet.setColumnWidth(cc, (cc < 16) ? 2550 : 2650);
		for (int a = 13; a < 18; a++)
			sheet.setColumnHidden(a, true);
		sheet1.setColumnWidth(0, 500);
		sheet1.setColumnWidth(1, 1800);
		sheet1.setColumnWidth(2, 8800);
		for (int cc = 3; cc < head1.size() + 1; cc++)
			sheet1.setColumnWidth(cc, 2700);
		sheet2.setColumnWidth(0, 500);
		sheet2.setColumnWidth(1, 1800);
		sheet2.setColumnWidth(2, 8800);
		for (int cc = 3; cc < head1.size() + 1; cc++)
			sheet1.setColumnWidth(cc, 2700);
		sheet.createFreezePane(5, 5);
		sheet1.createFreezePane(3, 4);
		sheet2.createFreezePane(3, 4);
		FileOutputStream xlsfileout = new FileOutputStream("d:/tnstore/review-" + ft.format(d) + ".xlsx");
		wb.write(xlsfileout);
		wb.close();
		Date end = new Date();
		long diff = end.getTime() - start.getTime();
		String TimeTaken = String.format("[%s] mins : [%s] secs", TimeUnit.MILLISECONDS.toMinutes(diff),
				TimeUnit.MILLISECONDS.toSeconds(diff));
		System.out.println(String.format("Completed ...... Time taken %s", TimeTaken));
	}
}
