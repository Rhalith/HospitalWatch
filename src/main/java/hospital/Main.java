package hospital;


import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) throws IOException {

        FileReader fileReader;
//        Scanner scanner = new Scanner(System.in);
//        System.out.println("Excel dosyasının bulunduğu klasör adresini gir.(Arama yerinin solundaki yere tıklayıp kopyala yapıştır)");
//        String adress = scanner.nextLine();
//        System.out.println("Excel Dosyasının adını gir.(f2'ye bas kopyala yapıştır)");
//        String excelName = scanner.nextLine();
//        System.out.println("Yeni oluşacak excel dosyasının adını gir.(Girmeyeceksen boş bırak)");
//        String newExcelName = scanner.nextLine();
//        System.out.println("Sayfa adını gir.(Örnek: EYLÜL 2023)");
//        String sheetName = scanner.nextLine();
//        System.out.println("Ayda kaç gün var?");
//        int days = scanner.nextInt();
        String adress = "C:\\Users\\nuhyi\\OneDrive\\Masaüstü\\Nöbet";
        String excelName = "nöbet";
        String sheetName = "EYLÜL 2023";
        int days = 30;
        List<Integer> skippingIndexes = new ArrayList<>();
        skippingIndexes.add(26);
        skippingIndexes.add(17);
        skippingIndexes.add(29);
        fileReader = new FileReader(adress, excelName,sheetName, days, skippingIndexes);
//        if(newExcelName.isEmpty()){
//            fileReader = new FileReader(adress, excelName,sheetName, days);
//        }
//        else{
//            fileReader = new FileReader(adress, excelName,sheetName, days, newExcelName);
//        }
        fileReader.readFile();
    }
}
