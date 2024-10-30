package src;

public class Test {
    //系统中ppt文件位置
    String inputFile ="C:\\Users\\Administrator\\Desktop\\Presentation.pptx";
    //输出文件的文件夹
    String outputFile="output";
    //创建一个ppt实例
    Presentation ppt = new Presentation();
//加载ppt文件
ppt.loadFromFile(inputFile);
//保存ppt文件为图像文件
for (int i = 0; i < ppt.getSlides().getCount(); i++) {
        BufferedImage image = ppt.getSlides().get(i).saveAsImage();
        String fileName = outputFile + "/" + String.format("ToImage-%1$s.png", i);
        ImageIO.write(image, "PNG",new File(fileName));
}
