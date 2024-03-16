<?php

require './vendor/autoload.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Background;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Slide\Background\Image;

class PhpOffice
{
    private const WIDTH = 1600;
    
    private const HEIGHT = 1000;
    
    private const LINE = '------------------------------------------------------------------------------';
    
    /**
     * 创建第一页或最后一页
     */ 
    public function createHeaderOrFooter($objPHPPowerPoint, string $title, string $imagepath, string $subtitle = '') 
    {
        $slide = $objPHPPowerPoint->createSlide();
        
        # 创建背景图
        $slide->setBackground((new Image())->setPath($imagepath));
        
        # 创建标题文本
        $this->createText($slide, $title, 60, self::WIDTH, 160, 250, Alignment::HORIZONTAL_LEFT)->setBold(true)->setSize(40)->setColor(new Color(Color::COLOR_BLACK));
        
        # 副标题文本
        if (!empty($subtitle)) {
            $this->createText($slide, $subtitle, 60, self::WIDTH, 160, 350, Alignment::HORIZONTAL_LEFT)->setBold(false)->setSize(25)->setColor(new Color('FF666666'));
        }
        
        # 创建小程序码
        $this->createImage($slide, './images/logo.png', 200, 200, 160, 550);
        
        # 创建说明文本
        $this->createText($slide, '扫码查看小程序', 60, 200, 160, 760)->setSize(16)->setColor(new Color('FF666666'));
    }
    
    /**
     * 绘制文本
     */ 
    private function createText($slide, string $text, int $height, int $width, int $x, int $y, string $position = Alignment::HORIZONTAL_CENTER) 
    {
        $shape = $slide->createRichTextShape()->setHeight($height)->setWidth($width)->setOffsetX($x)->setOffsetY($y);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal($position);
        $textRun = $shape->createTextRun()->setText($text);
        return $textRun->getFont();
    }
    
    /**
     * 绘制图片
     */ 
    private function createImage($slide, string $imagepath, int $height, int $width, int $x, int $y) 
    {
        $shape = $slide->createDrawingShape();
        $shape->setPath($imagepath)->setResizeProportional(false)->setHeight($height)->setWidth($width)->setOffsetX($x)->setOffsetY($y);
        return $shape;
    }
    
    /**
     * 绘制产品介绍文本
     */ 
    private function createProductText($slide, string $title, int $y, string $color = 'FF8C8C8C') {
        $this->createText($slide, $title, 40, 720, 740, $y, Alignment::HORIZONTAL_LEFT)->setSize(18)->setColor(new Color($color));
    }
    
    /**
     * 创建商品详情页
     */ 
    public function createProductDetail($slide, $mainImage, $images) 
    {
        # 左边商品图片组：1张大图，底下三个小图  
        $this->createImage($slide, $mainImage, 660, 660, 40, 40);
        
        # 大图下的3张小图
        $x = 40; $y = 720;
        foreach ($images as $image) {
            $this->createImage($slide, $image, 210, 210, $x, $y);
            $x += 225;
        }
        
        # 右侧文本
        $this->createProductText($slide, '商品ID：304513', 40 * 1);
        $this->createProductText($slide, '瑷露德玛 芦荟紧致滋养水120ml6935382613248', 40 * 2, 'FFB6712D');
        
        $this->createProductText($slide, self::LINE, 40 * 3, 'FFCCCCCC'); // 分隔符
        
        $this->createProductText($slide, '商品属性', 40 * 4, 'FFB6712D');
        $this->createProductText($slide, '商品毛重：140.00g', 40 * 5);
        $this->createProductText($slide, '商品产地：中国大陆', 40 * 6);
        $this->createProductText($slide, '特色功能：防汗', 40 * 7);
        $this->createProductText($slide, '佩戴方式：入耳式', 40 * 8);
        $this->createProductText($slide, '振膜类型：单动铁', 40 * 9);
        
        $this->createProductText($slide, '商品价格', 40 * 10, 'FFB6712D');
        $this->createProductText($slide, '装箱数：12   集采起订量：12', 40 * 11);
        $this->createProductText($slide, '集采价：¥140.4    电商价：¥149.00', 40 * 12);
        $this->createProductText($slide, '代发价：¥144.0 ', 40 * 13);
        
        $this->createProductText($slide, '商品卖点', 40 * 14, 'FFB6712D');
        $this->createProductText($slide, '①采用专利技术12小时完成库拉索芦荟从鲜叶到成品的加', 40 * 15);
        $this->createProductText($slide, '②芦荟精粹活性高，基础补水更尽兴，肌底渗透强', 40 * 16);
        $this->createProductText($slide, '③奢宠滋养呵护，肌肤整天都饱满莹润水嘟嘟', 40 * 17);
        $this->createProductText($slide, '④添加黄金胜肽肌肽，深层渗透，修护老化脆弱肌肤', 40 * 18);
        $this->createProductText($slide, '⑤活细胞元气，提升肌肤弹性，让肌肤饱满如婴儿肌', 40 * 19);
    }
    
    public function index()
    {
        // 1.创建ppt对象
        $objPHPPowerPoint = new PhpPresentation();
        
        # 2.自定义幻灯片尺寸
        $objPHPPowerPoint->getLayout()->setCX(self::WIDTH, DocumentLayout::UNIT_PIXEL)->setCY(self::HEIGHT, DocumentLayout::UNIT_PIXEL);

        // 3.设置属性
        $objPHPPowerPoint->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 02 Title')
            ->setSubject('Sample 02 Subject')
            ->setDescription('Sample 02 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        // 4.删除第一页(多页最好删除)
        $objPHPPowerPoint->removeSlideByIndex(0);
        
        # 第一页
        $this->createHeaderOrFooter($objPHPPowerPoint, '女神节心意礼品方案', './images/bg.jpeg');

        //根据需求 调整for循环
        for ($i = 1; $i <= 3; $i++) {
            
            //创建幻灯片并添加到这个演示中
            $slide = $objPHPPowerPoint->createSlide();
            
            # 构建产品介绍页
            $this->createProductDetail($slide, './images/1.png', ['./images/1.png', './images/2.png', './images/1.png']);
        }
        
        # 最后一页
        $this->createHeaderOrFooter($objPHPPowerPoint, '谢谢您的观看', './images/bg.jpeg');

        $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
        $url = './upload/' . time() . ".pptx";
        $oWriterPPTX->save($url);
        
        download($url);
    }
}

$class = new PhpOffice();
$class->index();
echo '----' . PHP_EOL;

function download($file)
{
    if(file_exists($file)){
        header("Content-type:application/octet-stream");
        $filename = basename($file);
        header("Content-Disposition:attachment;filename = ".$filename);
        header("Accept-ranges:bytes");
        header("Accept-length:".filesize($file));
        readfile($file);
    } else {
        echo "<script>alert('文件不存在')</script>";
    }
}

//删除文件
function deldir($dir)
{
    unlink($dir);
    closedir($dir);
}