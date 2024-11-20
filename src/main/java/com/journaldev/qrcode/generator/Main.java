package com.journaldev.qrcode.generator;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Hashtable;
import java.util.List;
import java.util.stream.Collectors;

import javax.imageio.ImageIO;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.WriterException;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.QRCodeWriter;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import java.nio.file.Paths;

public class Main {


	public static void main(String[] args) throws Exception {
		// Lấy đường dẫn của thư mục nơi file JAR đang chạy
		String jarDir = new File(Main.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getParent();

		// Đường dẫn file input
		String inputFilePath = Paths.get(jarDir, "qrcode_input.txt").toString();

		// Các đường dẫn khác (nếu cần, để cùng thư mục JAR)
		String templateFilePath = Paths.get(jarDir, "email_template.docx").toString();
		String outputDir = Paths.get(jarDir, "generated_emails").toString();
		String qrCodeDir = Paths.get(jarDir, "generated_qrcodes").toString();

		int qrCodeSize = 125;
		String qrCodeFileType = "png";

		// Tạo thư mục nếu chưa tồn tại
		new File(outputDir).mkdirs();
		new File(qrCodeDir).mkdirs();

		// Đọc file và xử lý như trước
		List<String> lines = readLinesFromFile(inputFilePath);
		for (String line : lines) {
			if (!line.trim().isEmpty()) {
				String[] var = line.split(",");
				String sanitizedFileName = var[1];
				String qrCodePath = Paths.get(qrCodeDir, sanitizedFileName + "." + qrCodeFileType).toString();
				generateQRCode(var[1], qrCodePath, qrCodeSize, qrCodeFileType);

				String outputFilePath = Paths.get(outputDir, sanitizedFileName + "_email.docx").toString();
				generateEmail(line, qrCodePath, templateFilePath, outputFilePath);

				System.out.println("Generated email template for: " + line);
			}
		}

		System.out.println("All templates generated successfully!");
	}

	private static List<String> readLinesFromFile(String filePath) throws IOException {
		try (BufferedReader reader = new BufferedReader(new FileReader(filePath))) {
			return reader.lines().collect(Collectors.toList());
		}
	}

	private static void generateQRCode(String text, String filePath, int size, String fileType)
			throws WriterException, IOException {
		Hashtable<EncodeHintType, ErrorCorrectionLevel> hintMap = new Hashtable<>();
		hintMap.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.L);

		QRCodeWriter qrCodeWriter = new QRCodeWriter();
		BitMatrix byteMatrix = qrCodeWriter.encode(text, BarcodeFormat.QR_CODE, size, size, hintMap);

		int matrixWidth = byteMatrix.getWidth();
		BufferedImage image = new BufferedImage(matrixWidth, matrixWidth, BufferedImage.TYPE_INT_RGB);
		image.createGraphics();

		Graphics2D graphics = (Graphics2D) image.getGraphics();
		graphics.setColor(Color.WHITE);
		graphics.fillRect(0, 0, matrixWidth, matrixWidth);
		graphics.setColor(Color.BLACK);

		for (int i = 0; i < matrixWidth; i++) {
			for (int j = 0; j < matrixWidth; j++) {
				if (byteMatrix.get(i, j)) {
					graphics.fillRect(i, j, 1, 1);
				}
			}
		}

		ImageIO.write(image, fileType, new File(filePath));
	}

	private static void generateEmail(String line, String qrCodePath, String templatePath, String outputFilePath) throws Exception {
		// Load the template
		try (FileInputStream fis = new FileInputStream(templatePath);
			 XWPFDocument document = new XWPFDocument(fis)) {

			// Replace placeholders
			for (XWPFParagraph paragraph : document.getParagraphs()) {
				for (XWPFRun run : paragraph.getRuns()) {
					String text = run.getText(0);

					String[] val = line.split(",");

					if (text != null && text.contains("{NAME}")) {
						text = text.replace("{NAME}", val[0]);
						run.setText(text, 0);
					}

					if (text != null && text.contains("{QR_CODE}")) {
						run.setText("", 0); // Clear the placeholder
						// Add QR code image
						try (FileInputStream qrStream = new FileInputStream(qrCodePath)) {
							run.addPicture(qrStream, Document.PICTURE_TYPE_PNG, qrCodePath, Units.toEMU(100), Units.toEMU(100));
						} catch (Exception e) {
							e.printStackTrace();
						}

					}
				}
			}

			// Save the updated document
			try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
				document.write(fos);
			}
		}
	}


}
