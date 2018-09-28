# This is a class that intents to read a DOCX file and output it to HTML format.

# USAGE

## debug mode
```
$rt = new WordPHP(true);
```

## without debug
```
$rt = new WordPHP(false); or $rt = new WordPHP();
```

## Set output encoding
```
$rt = new WordPHP(false, OUTPUT_ENCODING);
```

## Read docx file and returns the html code
```
$text = $rt->readDocument(FILENAME);
```

# NOTE:
To load images, please create a tmp folder where images can be extracted and saved from the document.
Don't forget to change the permissions in that folder so that the user that runs this class can write the images in it.
