package com.example.apache_poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Controller
public class excelTest {
    @GetMapping("/downloadExcel")
    public void downloadExcel(HttpServletResponse response) throws IOException {
        // 엑셀에 들어갈 데이터 생성
        product[] list = {
                new product(2022111501L, "노트북", 2000000),
                new product(2022111502L, "마우스", 130000),
                new product(2022111503L, "키보드", 230000),
                new product(2022111504L, "모니터", 820000)

        };

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("상품 리스트");
        int rowNo = 0;

        Row headerRow = sheet.createRow(rowNo++);
        headerRow.createCell(0).setCellValue("번호");
        headerRow.createCell(1).setCellValue("제품명");
        headerRow.createCell(2).setCellValue("금액");

        for (product s : list) {
            Row row = sheet.createRow(rowNo++);
            row.createCell(0).setCellValue(s.getProductId());
            row.createCell(1).setCellValue(s.getName());
            row.createCell(2).setCellValue(s.getAmount());
        }

        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=productList.xls");

        workbook.write(response.getOutputStream());
        workbook.close();

    }
}

class product {
    Long productId;
    String name;
    int amount;

    public product(Long productId, String name, int amount) {
        this.productId = productId;
        this.name = name;
        this.amount = amount;
    }

    @Override
    public String toString() {
        return "product [productId=" + productId + ", name=" + name + ", amount=" + amount + "]";
    }

    public Long getProductId() {
        return productId;
    }

    public void setProductId(Long productId) {
        this.productId = productId;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAmount() {
        return amount;
    }

    public void setAmount(int amount) {
        this.amount = amount;
    }
}
