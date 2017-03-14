package com.chinadaas.riskbell;

import java.io.*;

/**
 * Created by pc on 2017/3/13.
 */
public class ToHtml {

    public ToHtml(){

    }
    public static void main(String[] args){
        //输入类型判断
        BufferedReader bf=new BufferedReader(new InputStreamReader(System.in));
        System.out.println("请输入要进行的操作编号（1、2、3）：\n1、单个文件转换\n2、文件夹转换\n3、结束");
        String type="";
        while(true){
            try{
                type=bf.readLine();
                if(type.equals("1")==true){
                    System.out.println("你输入的是单个文件转换："+type);
                    //单个文件转换函数
                    System.out.print("请输入文件目录地址：");
                    BufferedReader b=new BufferedReader(new InputStreamReader(System.in));
                    String dirPath =b.readLine();
                    File file = new File(dirPath);
                    Els_x_toHtml els_x_toHtml=new Els_x_toHtml(dirPath,"F:\\"+file.getName()+".html");
                    System.out.println("转化后的输出路径为："+"F:\\"+file.getName()+".html");
                }
                else if(type.equals("2")==true){
                    System.out.println("你输入的是目录转换："+type);
                    //多文件转换接口
                    ToHtml toHtml=new ToHtml();

                    //循环调用单个文件转换
                    System.out.print("请输入文件目录地址：");
                    BufferedReader bf2=new BufferedReader(new InputStreamReader(System.in));
                    String dirPath2 =bf2.readLine();
                    System.out.print("请输入文件输出目录地址：");
                    BufferedReader bf3=new BufferedReader(new InputStreamReader(System.in));
                    String dirPath3 =bf3.readLine();
                    Dir_todir dir_todir = new Dir_todir(dirPath2,dirPath3);

                }
                else if(type.equals("3")==true){
                    System.out.println("结束"+type);
                    break;
                }
                else{
                    System.out.println("你输入的不是有效内容");
                    System.out.println("请输入要进行的操作编号（1、2、3）：\n1、单个文件转换\n2、文件夹转换\n3、结束");
                }
            } catch (IOException e) {   e.printStackTrace();}
        }
        System.out.print("done");
    }

}
