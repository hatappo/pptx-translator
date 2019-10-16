# ppt-tr

A sample code.

Traverse objects in a PowerPoint file which has `pptx` extension, extract all (almost) text elements, translate using Amazon Translate, and output as a new file.

## Usage

In the default, a target input file is `resources/test.pptx` and output file is `resources/test.pptx.translated-at-{yyyyMMddHHMMss}.pptx`
You can change it because it's written in the source code directly.

## TODO

* clean
* translate "Note" objects.
* lgging for processing time.
* ...

## Reference

### reference for POI
* https://www.tutorialspoint.com/apache_poi_ppt/apache_poi_ppt_formatting_text.htm
* https://qiita.com/Lain_/items/6a1d3ed0255720d6d57a
* https://gist.github.com/ponkore/4216377
* https://poi.apache.org/apidocs/dev/org/apache/poi/xslf/usermodel/XSLFTextBox.html

### reference for Amazon Translate and Java SDK
* https://docs.aws.amazon.com/translate/latest/dg/examples-java.html
