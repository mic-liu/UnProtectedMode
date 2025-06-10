package com.example;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.*;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import java.util.Scanner;

public class XlsmProcessor {
    private static boolean overwriteOriginal = false;

    public static void main(String[] args) {
        if (args.length == 0) {
            System.out.println("Excelファイルまたはフォルダをこのプログラムにドラッグ＆ドロップしてください！");
            return;
        }

        System.out.println("\n");
        System.out.println("************************************************");
        System.out.println("ToolName：CCUSExcelDoc保護モード解除ツール");
        System.out.println("  Author：Mic.Liu");
        System.out.println(" Version：v1.0.1");
        System.out.println("************************************************");
        // ユーザーに処理方法を確認
        askProcessingOption();

        // ドラッグ＆ドロップされた各ファイルまたはフォルダを処理
        for (String path : args) {
            processPath(new File(path));
        }

        System.out.println("すべてのファイルの処理が完了しました！");
        System.out.println("任意のキーを押して終了...");
        try {
            System.in.read();
        } catch (IOException e) {
            // 例外を無視
        }
    }

    private static void askProcessingOption() {
        Scanner scanner = new Scanner(System.in);
        while (true) {
            System.out.println("ファイルの処理方法を選択してください：");
            System.out.println("--------------------------------------------------------");
            System.out.println("0. 処理を終了");
            System.out.println("1. 新しいファイルを生成（ファイル名に_unprotectedを付加）");
            System.out.println("2. 元のファイルを直接上書き");
            System.out.println("--------------------------------------------------------");
            System.out.print("オプションを入力してください（0、1または2）：");

            String input = scanner.nextLine().trim();
            if (input.equals("0")) {
                System.out.println("処理を終了します。");
                System.exit(0); 
            } else if (input.equals("1")) {
                overwriteOriginal = false;
                break;
            } else if (input.equals("2")) {
                overwriteOriginal = true;
                break;
            } else {
                System.out.println("\n無効なオプションです。もう一度選択してください。\n");
            }
        }
        System.out.println(); // 空行を印刷
    }

    private static void processPath(File path) {
        if (path.isDirectory()) {
            // フォルダを処理
            System.out.println("フォルダを処理中：" + path.getAbsolutePath());
            File[] files = path.listFiles();
            if (files != null) {
                for (File file : files) {
                    processPath(file);
                }
            }
        } else {
            // 単一ファイルを処理
            processFile(path.getAbsolutePath());
        }
    }

    private static void processFile(String inputFile) {
        String lowerInputFile = inputFile.toLowerCase();
        if (!lowerInputFile.endsWith(".xlsm") && 
            !lowerInputFile.endsWith(".xlsx") && 
            !lowerInputFile.endsWith(".xls")) {
            return;
        }

        // ユーザーの選択に基づいて出力ファイルパスを決定
        String outputFile;
        if (overwriteOriginal) {
            outputFile = inputFile;
        } else {
            outputFile = inputFile.substring(0, inputFile.lastIndexOf(".")) + 
                         "_unprotected" + 
                         inputFile.substring(inputFile.lastIndexOf("."));
        }

        try {
            // 各ファイルに一意の一時ディレクトリを作成
            File tempDir = new File("temp_" + System.currentTimeMillis());
            tempDir.mkdir();

            System.out.println("処理中：" + inputFile);

            // 上書きモードの場合、バックアップを作成
            if (overwriteOriginal) {
                String backupFile = inputFile + ".bak";
                Files.copy(Paths.get(inputFile), Paths.get(backupFile), StandardCopyOption.REPLACE_EXISTING);
            }

            // Excelファイルを解凍
            unzipFile(inputFile, tempDir.getPath());

            // すべてのXMLファイルを処理
            processXmlFiles(tempDir);

            // 再パッケージ化
            zipDirectory(tempDir, outputFile);

            // 一時ファイルをクリーンアップ
            deleteDirectory(tempDir);

            // 上書きモードの場合、処理成功後にバックアップを削除
            if (overwriteOriginal) {
                Files.delete(Paths.get(inputFile + ".bak"));
            }

            System.out.println("処理完了！" + (overwriteOriginal ? "元のファイルを上書きしました" : "出力ファイル：" + outputFile));
        } catch (Exception e) {
            System.out.println("処理失敗：" + inputFile);
            // 上書きモードの場合、エラー発生時にバックアップを復元
            if (overwriteOriginal) {
                try {
                    Files.move(Paths.get(inputFile + ".bak"), Paths.get(inputFile), 
                              StandardCopyOption.REPLACE_EXISTING);
                    System.out.println("元のファイルを復元しました");
                } catch (IOException restoreError) {
                    System.out.println("元のファイルの復元に失敗しました！バックアップファイル：" + inputFile + ".bak");
                }
            }
            e.printStackTrace();
        }
    }

    private static void unzipFile(String zipFile, String outputFolder) throws IOException {
        byte[] buffer = new byte[1024];
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile))) {
            ZipEntry zipEntry = zis.getNextEntry();
            while (zipEntry != null) {
                String fileName = zipEntry.getName();
                File newFile = new File(outputFolder + File.separator + fileName);
                
                // 创建目录
                new File(newFile.getParent()).mkdirs();

                if (!zipEntry.isDirectory()) {
                    try (FileOutputStream fos = new FileOutputStream(newFile)) {
                        int len;
                        while ((len = zis.read(buffer)) > 0) {
                            fos.write(buffer, 0, len);
                        }
                    }
                }
                zipEntry = zis.getNextEntry();
            }
        }
    }

    private static void processXmlFiles(File directory) throws IOException {
        // 添加workbookProtection的正则表达式
        Pattern sheetProtectionPattern = Pattern.compile("<sheetProtection\\b[^>]*\\/>");
        Pattern workbookProtectionPattern = Pattern.compile("<workbookProtection\\b[^>]*\\/>");
        
        if (directory.isDirectory()) {
            for (File file : directory.listFiles()) {
                if (file.isDirectory()) {
                    processXmlFiles(file);
                } else if (file.getName().endsWith(".xml")) {
                    // 读取文件内容
                    String content = readFile(file);
                    // 替换匹配的内容
                    content = sheetProtectionPattern.matcher(content).replaceAll("");
                    content = workbookProtectionPattern.matcher(content).replaceAll("");
                    // 写回文件
                    writeFile(file, content);
                }
            }
        }
    }

    private static String readFile(File file) throws IOException {
        String encoding = "UTF-8";

        try (FileInputStream fis = new FileInputStream(file);
             BufferedInputStream bis = new BufferedInputStream(fis)) {
            
            bis.mark(4);
            byte[] bom = new byte[4];
            int n = bis.read(bom, 0, 4);
            
            if (n >= 3 && bom[0] == (byte)0xEF && bom[1] == (byte)0xBB && bom[2] == (byte)0xBF) {
                encoding = "UTF-8";
                bis.reset();
                bis.skip(3);
            } else if (n >= 2 && bom[0] == (byte)0xFF && bom[1] == (byte)0xFE) {
                encoding = "UTF-16LE";
                bis.reset();
                bis.skip(2);
            } else if (n >= 2 && bom[0] == (byte)0xFE && bom[1] == (byte)0xFF) {
                encoding = "UTF-16BE";
                bis.reset();
                bis.skip(2);
            } else {
                bis.reset();
            }
            
            StringBuilder content = new StringBuilder();
            char[] buffer = new char[8192];
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(bis, encoding))) {
                int len;
                while ((len = reader.read(buffer)) != -1) {
                    content.append(buffer, 0, len);
                }
            }
            return content.toString();
        }
    }

    private static void writeFile(File file, String content) throws IOException {
        // First read the file to detect original encoding
        String encoding = "UTF-8";
        byte[] bom = null;

        try (FileInputStream fis = new FileInputStream(file);
             BufferedInputStream bis = new BufferedInputStream(fis)) {
            
            bis.mark(4);
            byte[] bomCheck = new byte[4];
            int n = bis.read(bomCheck, 0, 4);
            
            if (n >= 3 && bomCheck[0] == (byte)0xEF && bomCheck[1] == (byte)0xBB && bomCheck[2] == (byte)0xBF) {
                encoding = "UTF-8";
                bom = new byte[]{(byte)0xEF, (byte)0xBB, (byte)0xBF};
            } else if (n >= 2 && bomCheck[0] == (byte)0xFF && bomCheck[1] == (byte)0xFE) {
                encoding = "UTF-16LE";
                bom = new byte[]{(byte)0xFF, (byte)0xFE};
            } else if (n >= 2 && bomCheck[0] == (byte)0xFE && bomCheck[1] == (byte)0xFF) {
                encoding = "UTF-16BE";
                bom = new byte[]{(byte)0xFE, (byte)0xFF};
            }
        }

        try (FileOutputStream fos = new FileOutputStream(file)) {
            // Write BOM if it existed in the original file
            if (bom != null) {
                fos.write(bom);
            }
            
            // Write content with detected encoding
            try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(fos, encoding))) {
                writer.write(content);
                writer.flush();
            }
        }
    }

    private static void zipDirectory(File directory, String zipFile) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(zipFile);
             ZipOutputStream zos = new ZipOutputStream(fos)) {
            zipFile(directory, "", zos);
        }
    }

    private static void zipFile(File directory, String baseName, ZipOutputStream zos) throws IOException {
        File[] files = directory.listFiles();
        byte[] buffer = new byte[1024];
        
        for (File file : files) {
            String entryName = baseName + (baseName.isEmpty() ? "" : File.separator) + file.getName();
            
            if (file.isDirectory()) {
                zipFile(file, entryName, zos);
                continue;
            }
            
            ZipEntry ze = new ZipEntry(entryName);
            zos.putNextEntry(ze);
            
            try (FileInputStream in = new FileInputStream(file)) {
                int len;
                while ((len = in.read(buffer)) > 0) {
                    zos.write(buffer, 0, len);
                }
            }
        }
    }

    private static void deleteDirectory(File directory) {
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    deleteDirectory(file);
                } else {
                    file.delete();
                }
            }
        }
        directory.delete();
    }
}