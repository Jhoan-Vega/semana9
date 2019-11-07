/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package semana09;
import java.io.*;
import javax.swing.*;
import jxl.write.*;
import jxl.*;
/**
 *
 * @author joan_
 */
public class ExportarExcel {
    private File file;
    private JTable table;
    private String nombreTab;
    
    public ExportarExcel(JTable table,File file,String nombreTab){
        this.file=file;
        this.table=table;
        this.nombreTab=nombreTab;
    }
    
    public boolean export(){
        try
        {
            System.out.println("iniciando proceso de exportar");
            DataOutputStream out=new DataOutputStream(new FileOutputStream(file));
            WritableWorkbook w= Workbook.createWorkbook(out);
            System.out.println("colocando nombre");
            WritableSheet s=w.createSheet(nombreTab, 0);
            System.out.println("recorriendo la tabla");
            for (int i = 0; i < table.getRowCount(); i++) {
                for (int j = 0; j < table.getColumnCount(); j++) {
                    Object objeto=table.getValueAt(i,j);
                    s.addCell(new Label(j,i, String.valueOf(objeto)));
                }
            }
            w.write();
            System.out.println("cerramos el writable y el data");
            w.close();
            out.close();
            return true;
        }catch(IOException ex){ex.printStackTrace();}
        catch(WriteException ex){ex.printStackTrace();}
        System.out.println("ocurrio un error");
        return false;
    }
}
