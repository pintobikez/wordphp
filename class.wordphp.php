<?php

class WordPHP
{
	private $debug = false;
	private $rels_xml;
	private $doc_xml;
	private $description='';
	private $d_spotlight='';
	private $d_webpage_interest='';
	private $d_bibliography_documental='';
	private $d_bibliography_fiction='';
	private $d_audiovisual='';
	private $practical_advices='';
	private $last = 'none';
	
	public function __construct($debug_=null){
		if($debug_ != null)
			$this->debug = $debug_;
	}
	
	private function readZipPart($filename){
		$zip = new ZipArchive();
		$_xml = 'word/document.xml';
		$_xml_rels = 'word/_rels/document.xml.rels';
		
		if (true === $zip->open($filename)) {
			if (($index = $zip->locateName($_xml)) !== false) {
				$xml = $zip->getFromIndex($index);
			}
			$zip->close();
		} else die('non zip file');
		
		if (true === $zip->open($filename)) {
			if (($index = $zip->locateName($_xml_rels)) !== false) {
				$xml_rels = $zip->getFromIndex($index);					
			}
			$zip->close();
		} else die('non zip file');
		
		$this->doc_xml = new DOMDocument();
		$this->doc_xml->encoding = mb_detect_encoding($xml);
		$this->doc_xml->preserveWhiteSpace = false;
		$this->doc_xml->formatOutput = true;
		$this->doc_xml->loadXML($xml);
		$this->doc_xml->saveXML();
		
		$this->rels_xml = new DOMDocument();
		$this->doc_xml->encoding = mb_detect_encoding($xml);
		$this->rels_xml->preserveWhiteSpace = false;
		$this->rels_xml->formatOutput = true;
		$this->rels_xml->loadXML($xml_rels);
		$this->rels_xml->saveXML();
		
		if($this->debug){
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->doc_xml->saveXML();
			echo "</textarea>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->rels_xml->saveXML();
			echo "</textarea>";
		}
	}

	private function checkFormating(&$formatting, &$xml)
	{	
		$node = trim($xml->readOuterXML());		
		// add <br> tags
		if (strstr($node,'<w:br ')) $text .= '<br>';					 
		// look for formatting tags
		$f = "<span style='";
		$reader = new XMLReader();
		$reader->XML($node);
		while ($reader->read())
		{
			if($reader->name == "w:b")
				$f .= "font-weight: bold,";
			if($reader->name == "w:i")
				$f .= "text-decoration: underline,";
			if($reader->name == "w:color")
				$f .="color: #".$reader->getAttribute("w:val").",";
			if($reader->name == "w:rFont")
				$f .="font-family: #".$reader->getAttribute("w:ascii").",";
			if($reader->name == "w:shd" && $reader->getAttribute("w:val") != "clear" && $reader->getAttribute("w:fill") != "000000")
				$f .="background-color: #".$reader->getAttribute("w:fill").",";
		}
		
		$f = rtrim($f, ',');
		$f .= "'>";
		
		return $f.htmlentities($xml->expand()->textContent)."</span>";
	}
	
	private function getListFormating(&$xml)
	{	
		$node = trim($xml->readOuterXML());
		
		$reader = new XMLReader();
		$reader->XML($node);
		$ret="";
		$close = "";
		while ($reader->read())
		{
			if($reader->name == "w:numPr" && $reader->nodeType == XMLReader::ELEMENT )
			{
						
			}
			if($reader->name == "w:numId" && $reader->hasAttributes)
			{
				switch($reader->getAttribute("w:val"))
				{
					case 1:
						$ret['open'] = "<ol><li>";
						$ret['close'] = "</li></ol>";
						break;
					case 2:
						$ret['open'] = "<ul><li>";
						$ret['close'] = "</li></ul>";
						break;
				}
				
			}
		}
	
		return $ret;
	}
	
	private function checkImageFormating(&$xml){
		
	}
	
	private function getHyperlink(&$xml)
	{
		$ret = array('open'=>'<ul>','close'=>'</ul>');
		$link ='';
		if($xml->hasAttributes)
		{
			$attribute = "";
			while($xml->moveToNextAttribute()) 
			{
				if($xml->name == "r:id")
					$attribute = $xml->value;
			}
			
			if($attribute != "")
			{
				$reader = new XMLReader();
				$reader->XML($this->rels_xml->saveXML());
				
				while ($reader->read())
				{
					if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name=='Relationship')
					{
						if($reader->getAttribute("Id") == $attribute){
							$link = $reader->getAttribute('Target');
							break;
						}
					}
				}
			}
		}
		
		if($link != "")
		{
			$ret['open'] = "<a href='".$link."' target='_blank'>";
			$ret['close'] = "</a>";
		}
		
		return $ret;
	}
	
	public function readDocument($filename) {
		
		$this->readZipPart($filename);
		$reader = new XMLReader();
		$reader->XML($this->doc_xml->saveXML());

		$text = ''; $list_format="";

		$formatting['header'] = 0;
		// loop through docx xml dom
		while ($reader->read()){
		// look for new paragraphs
			$paragraph = new XMLReader;
			$p = $reader->readOuterXML();
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:p'){
				// set up new instance of XMLReader for parsing paragraph independantly				
				$paragraph->xml($p);

				preg_match('/<w:pStyle w:val="(Heading.*?[1-6])"/',$p,$matches);
				if(isset($matches[1])){
				switch($matches[1]){
				case 'Heading1': $formatting['header'] = 1; break;
				case 'Heading2': $formatting['header'] = 2; break;
				case 'Heading3': $formatting['header'] = 3; break;
				case 'Heading4': $formatting['header'] = 4; break;
				case 'Heading5': $formatting['header'] = 5; break;
				case 'Heading6': $formatting['header'] = 6; break;
				default: $formatting['header'] = 0; break;
				}}
				// open h-tag or paragraph
				$text .= ($formatting['header'] > 0) ? '<h'.$formatting['header'].'>' : '<p>';
				
				// loop through paragraph dom
				while ($paragraph->read()){
					// look for elements
					if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:r'){
						if($list_format == "")
							$text .= $this->checkFormating($formatting, $paragraph);
						else
						{
							$text .= $list_format['open'];
							$text .= $this->checkFormating($formatting, $paragraph);
							$text .= $list_format['close'];
						}
						$list_format ="";
						$paragraph->next();
					}
					else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:pPr'){ //lists
						$list_format = $this->getListFormating($paragraph);
						$paragraph->next();
					}
					else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:drawing'){ //images
						$text .= $this->checkImageFormating($paragraph);
						$paragraph->next();
					}
					else if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:hyperlink'){
						$hyperlink = $this->getHyperlink($paragraph);
						$text .= $hyperlink['open'];
						$text .= $this->checkFormating($formatting, $paragraph);
						$text .= $hyperlink['close'];
						$paragraph->next();
					}
				}
				$text .= ($formatting['header'] > 0) ? '</h'.$formatting['header'].'>' : '</p>';
			}
		}
		$reader->close();
		if($this->debug){
			echo "<div style='width:100%; height: 200px;'>";
			echo iconv("ISO-8859-1", "UTF-8",$text);
			echo "</div>";
		}
		return iconv("ISO-8859-1", "UTF-8",$text);
	}
}