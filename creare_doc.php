<?php 
class word 
{  

	function test()
	{
		echo 'hello world';
	}
	function start() 
	{ 
	ob_start(); 
	print'<html xmlns:o="urn:schemas-microsoft-com:office:office" 
	xmlns:w="urn:schemas-microsoft-com:office:word" 
	xmlns="http://www.w3.org/TR/REC-html40">'; 

	} 

	function save($path) 
	{ 

	print "</html>"; 
	$data = ob_get_contents(); 

	ob_end_clean(); 

	$this->wirtefile ($path,$data); 
	} 

	function wirtefile ($fn,$data) 
	{ 

	$fp=fopen($fn,"wb"); 
	fwrite($fp,$data); 
	fclose($fp); 
	} 

} 



$word=new word; 

$word->start(); 
?> 

<title>直接用php创建word文档</title> 
 <h1>直接用php创建word文档</h1> 
 作者：axgle 
<hr size=1> 
 <p>如果你打开word.doc，看到了这里的介绍，则说明word文档创建成功了。 
<p> 
不论是在什么操作系统下，使用本方法都可以直接用PHP生成word文档。绝对不是吹牛！ 
就算是没有安装word，也能够生成word文件。 
当然了，生成的word文件可以用word,wps或者其他软件打开。 
<p> 
<b>使用方法：</b> 
<br> 
首先用$word->start()表示要生成word文件了。 
然后你可以输出任何的HTML代码，不论是从文件读过来再写到这里， 
还是直接在这里输出HTML，都没有关系。 

<p>等你输出完毕后，用$word->save($path)方法，其中$path是你想 
生成的word文件的名称（可以给出完整的路径）.当你使用了$word->save() 
方法后，这后面的任何输出都和word文件没有关系了，也就是说word的生成 
工作就完成了。之后就和你平常使用php的方式一样拉。随便你输出什么东西， 
都直接在浏览器里输出，而不会写到word里面去。 
<p>这是本人想到的一个很有意思的方法，它的实现方法出人意料的简单，并且避免 
了对windows环境的依赖。 
<br>哈哈，很有意思吧？享受它吧！ 
<hr size=1> 
<?php
$word->save("data.doc");//保存word并且结束. 

echo ' 
<title>直接用php创建word文档</title> 
<h1>直接用php创建word文档</h1> 
生成word了吗？在你的目录下看看有没有data.doc! 
<br> 
如果你用的是windows，并且安装得有word,可以查看<p> 
<a href="data.doc" target=_blank>这里</a>'; 

?>
