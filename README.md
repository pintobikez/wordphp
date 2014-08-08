This is a class that intents to read a DOCX file and output it to HTML format.

USAGE

//debug mode
$rt = new WordPHP(true);

//without debug
$rt = new WordPHP(false); or $rt = new WordPHP();

//Read docx file and returns the html code
$text = $rt->readDocument(FILENAME);
