=== Word to html ===
Contributors: wibergsweb
Tags: docx, convert, converter, word to html, word into html, convert, external, multiple
Requires at least: 3.0.1
Requires PHP: 5.2.4
Tested up to: 5.2
Stable tag: 1.1
License: GPLv2
License URI: http://www.gnu.org/licenses/gpl-2.0.html

Display some html from one or more word files from your local webserver or an external webserver. 

== Description ==

Word to html makes it easy to fetch content from newer word-file(s) (.docx) and convert the document to html on a page with a single shortcode.  

Display some html from one or more word files from your local webserver or an external webserver. Multiple word files support on your local webserver. It's not intented to be used as full format converting tool as it does create clean html (not a loads of inline styles), but could be convient to read content from word-files just by putting them in a folder on your webserver.

The plugin does fetch this kind of information from your .docx - document:

* Headings (support for english and swedish)
* Paragraphs including bold, underline and italic. When having paragraph defined as a column then a new paragraph is created.
* Hyperlinks
* Tables 
* Images (jpg, png, jpeg)
* Unordered lists (bulleted) or ordered lists (numbered). It's also possible to combine these type of lists.

The plugin does only create the html without any specific inline styling. The reason for this is that word-documents and html-documents have a totally different structure and you should be using css when styling html - documents. Inline styles would also apply styling to your wordpress-theme that wouldn't fit your current design.

If fetching information from more then one docx - document on your local webserver, content from all files are mixed into different sections of html (one file per section). This could be useful if you want to create tabs of some sort based on information from several word-documents. 

If you like the plugin, please consider donating or write a review.

== Installation ==

This section describes how to install the plugin and get it working.

1. Upload the plugin folder wordtohtml to the `/wp-content/plugins/' directory, or install the plugin through the WordPress plugins screen directly.
2. Activate the plugin through the 'Plugins' screen in WordPress
3. Put shortcode on the Wordpress post or page you want to display it on and add css to change layout for those.

= Shortcodes =
* [wordtohtml_create] - Create the html from specified word-file(s)

= [wordtohtml_create] attributes =
* html_id - set id of the article elemenet (structure i article/section/section.../article)
* html_class - set class of section elements
* path - relative path to uploads-folder of the wordpress - installation ( eg. /wp-content/uploads/{path} ). If fetching docx documents from an external site then that document is copied to this path and after that the file is being processed locally.
* source_files - file(s) to include. If using more than one file - separate them with a semicolon (;). It 's possible to include a full url instead of a filename to fetch external word files. It's also possible to fetch all files from a given path (with for example *). 
* eol_detection - CR = Carriage return, LF = Line feed, CR/LF = Carriage line and line feed, auto = autodetect. Only useful on external files. Other files are automatically autodeteced.
* convert_encoding_from - When converting character encoding, define what current characterencoding that word file has. (Not required, but gives best result)
* convert_encoding_to - When converting character encoding, define what characterencoding that word should be encoded to. (Best result of encoding is when you define both encoding from and encoding both)
* add_ext_auto - Add fileextension .docx to file. Set to no if you don't want the file extension to be added automatically.
* skip_articletag - Default is set to yes (probably you won't need this because wordpress does have an article-tag already in the document)
* debug_mode - if set to yes then then important debugging information will be displayed (otherwise it would be "silent errors")

= Default values =
* [wordtohtml_create html_id="{none}" html_class="{none}" path="{none}" source_files="{none}" eol_detection="auto" convert_encoding_from="{none}" convert_encoding_to="{to}" add_ext_auto="yes" skip_articletag="yes" debug_mode="no"]



== Frequently Asked Questions ==

= Why don't you include any css for the plugin? =

The goal is to make the plugin work as fast as possible as expected even with the design. By not supplying any css the developer has full control over
the design. There is actually one row of css but that does hardly count.



== Screenshots ==

No screenshots available

== Changelog ==

= 1.0 =
* Plugin released

= 1.1 =
* Tables bugfix. When there are serveal tables then they are located at correct location in html document. Colspan and table headers are applied.
* Bugfixes headers
* Bugfixes lists (unordered and ordered list are both applied and can be combined)
* Bugfixes images (include only valid imagetypes into the html)
* Bugfixes Hyperlinks (urls and mailto) applied correctly.
* It's now possible to skip articletag
* It's now possible to load external documents (docx) and load it into html.
* Clean html (use css to style)
* Debugging texts improved

== Upgrade notice ==
If you're using 1.0 and upgrading to 1.1 then you should be aware that some things are removed to make the html much cleaner. This means that the word-document 
might fail to render as the word-document shows. The reason to this is that is merely impossible to make a such feature impossible. It's better to style the finished 
document with css (yourself or a developer that may help you)

Please tell me if you're missing something (in the support form) ! I will do my best to add the feature.

== Example of usage ==

= shortcodes in post(s)/page(s) =
* [wordtohtml_create path="lan" source_files="skane.docx;smaland.docx;"]
* [wordtohtml_create path="wordfiles" source_files="*.docx"]
* [wordtohtml_create path="wordfiles" source_files="word1;word2" debug_mode="yes"]
* [wordtohtml_create path="wordfiles" source_files="word1;word2" debug_mode="yes"]
* [wordtohtml_create debug_mode="no" convert_encoding_from="Windows-1252" convert_encoding_to="UTF-8"]
* [wordtohtml_create debug_mode="no" convert_encoding_to="UTF-8"]
* [wordtohtml_create source_files="http://wibergsweb.se/konstak.docx" path="wordfiles" debug_mode="yes" html_id="turnover" html_class="wow" add_ext_auto="yes" skip_articletag="no"]