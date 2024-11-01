<?php
/*
Plugin Name: Word to html
Plugin URI: http://www.wibergsweb.se/plugins/wordtohtml
Description:From one or more word files - create clean html.
Version: 1.1
Author: Wibergs Web
Author URI: http://www.wibergsweb.se/
Text Domain: wordtohtml-wp
Domain Path: /lang
License: GPLv2
*/
defined( 'ABSPATH' ) or die( 'No access allowed!' );

if( !class_exists('wordtohtmlwp') ) {
    ini_set("auto_detect_line_endings", true); //Does not apply when loading external file(s), therefore also custom function for this below


    //Main class
    class wordtohtmlwp
    {
    public $errormessage;
    private $default_eol = "\r\n"; //Default - use this as this has been default in previous version of the plugin
    private $encoding_to = null;
    private $encoding_from = null;
    
    /*
    *  Constructor
    *
    *  This function will construct all the neccessary actions, filters and functions for the sourcetotable plugin to work
    *
    *
    *  @param	N/A
    *  @return	N/A
    */	
    public function __construct() 
    {                        
        add_action( 'init', array( $this, 'loadlanguage' ) );
    }
    

    /*
     * loadjs
     * 
     * This function load javascript and (if there are any) translations
     *  
     *  @param	N/A
     *  @return	N/A
     *                 
     */    
    public function loadlanguage() 
    {                       
        //Load (if there are any) translations
        $loaded_translation = load_plugin_textdomain( 'wordtohtml-wp', false, dirname( plugin_basename(__FILE__) ) . '/lang/' );
        
        wp_enqueue_style(
            'wthtml_css',
            plugins_url( '/css/wibergsweb.css' , __FILE__ )
        );     
                
        $this->init();
   }      


    /*
     *  error_notice
     * 
     *  This function is used for handling administration notices when user has done something wrong when initiating this object
     *  Shortcode-equal to: No shortcode equavilent
     * 
     *  @param N/A
     *  @return N/A
     *                 
     */                 
    public function error_notice() 
    {
        $message = $this->errormessage;
        echo __("<pre><strong>Word to html Error:</strong><p>{$message}</p></pre>", 'wordtohtml-wp');
    }


    /*
     *  init
     * 
     *  This function initiates the actual shortcodes etc
     * 
     *  @param N/A
     *  @return N/A
     *                 
     */        
    public function init() 
    {               
        //Add shortcodes
        add_shortcode( 'wordtohtml_create', array ( $this, 'source_to_html') );
        add_action( 'admin_menu', array( $this, 'help_page') );
    }
    
    public function help_page() 
    {
        add_management_page( 'Word to html', 'Word to html', 'manage_options', 'word-to-html', array( $this, 'startpage') );                                 
    }
    
    
    /*
     *  init
     * 
     *  This function initiates the actual shortcodes etc
     * 
     *  @param N/A
     *  @return N/A
     *                 
     */    
    public function startpage() 
    {
      $html = '<h1>Word to html</h1>';
      $html .= '<div class="wrap">';
      $html .= '<ul>';
      $html .= '<li>';
      $html .= __('Word to html is a plugin that converts one or more word-files (docx) and outputs them as html.', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<li>';
      $html .= __('If you change something in the word-file(s) this would be reflected in the output of the html directly.', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<li>';
      $html .= __('To get this to work you need to create a shortcode and you need at least one word-file.', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '</ul>';
      
      $html .= '<ul>';
      $html .= '<li>';
      $html .= '<h2>' . __('Shortcode', 'wordtohtml-wp') . '</h2>';
      $html .= __('A shortcode in wordpress is basically just a placeholder that tells some code to execute. A shortcode for this plugin tells some of this plugins code to execute', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<li>';
      $html .= __('Example: [<strong>wordtohtml_create</strong> html_class=”wordtohtml” source_files=”sweden.docx” path="maps"]', 'wordtohtml-wp');
      $html .= ' The shortcode is [<strong>wordtohtml_create</strong>] and the shortcode-attributes are html_class, source_files and path. You may put a shortcode anywhere: in a post, in a page, in a widget, or even in a template.';
      $html .= '</ul>';

      $html .= '<ul>';
      $html .= '<li>';
      $html .= '<h2>' . __('Word files', 'wordtohtml-wp') . '</h2>';
      $html .= __('A word-file of the format .docx is based on xml-structure. In fact, one .docx document is zip-archive that contains serveral different xml-files.', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<li>';
      $html .= __('The Word to html plugin can handle reading files locally on the webbserver or from a remote location (through the http/https-protocol)', 'wordtohtml-wp');
      $html .= '</ul>';

      $html .= '<ul>';
      $html .= '<li>';
      $html .= '<h2>' . __('Reading word files locally', 'wordtohtml-wp') . '</h2>';
      $html .= __('To read word-file locally (on same webbserver that wordpress exists on) you have to define at least one csv-file. It\'s important to understand that the plugin has it\'s root from the uploads-folder of your wordpress (Look below in examples for clarification).', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<hr>';
      $html .= '<li>';
      $html .= __('<i>Example 1:</i> <strong>[wordtohtml_create source_files=”sweden.docx”]</strong>', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<li>';
      $upload_dir = wp_upload_dir();
      $upload_basedir = $upload_dir['basedir'];
        
      $html .= __('This example would make create html from the file sweden.docx that exists in', 'wordtohtml-wp') . ' ' . $upload_basedir . '';
      $html .= '</li>';
      $html .= '<hr>';
      $html .= '<li>';
      $html .= __('<i>Example 2:</i> <strong>[wordtohtml_create source_files=”sweden.docx;norway.docx;iceland.docx”]</strong>', 'wordtohtml-wp');
      $html .= '</li>';    
      $html .= '<li>';
      $html .= __('This example would make create html from the files sweden.docx, norway.docx and iceland.docx that exists in', 'wordtohtml-wp') . ' ' . $upload_basedir . '';
      $html .= '</li>';
      $html .= '<hr>';
  
      $html .= '<li>';
      $html .= __('<i>Example 3:</i> <strong>[wordtohtml_create path="mapfiles" source_files=”sweden.docx;norway.docx;iceland.docx”]</strong>', 'wordtohtml-wp');
      $html .= '</li>';
    
      $html .= '<li>';
      $html .= __('This example would make an html table from the files sweden.docx, norway.docx and iceland.docx that exists in', 'wordtohtml-wp') . ' ' . $upload_basedir . '/mapfiles';
      $html .= '</li>';
      $html .= '<hr>';

      $html .= '<li>';
      $html .= __('<i>Example 4:</i> <strong>[wordtohtml_create path="mapfiles" source_files=”*.docx”]</strong>', 'wordtohtml-wp');
      $html .= '</li>';
    
      $html .= '<li>';
      $html .= __('This example would make an html table from all the files that exists in', 'wordtohtml-wp') . ' ' . $upload_basedir . '/mapfiles';
      $html .= '</li>';
      $html .= '<hr>';
      
      $html .= '</ul>';
      
    $html .= '<ul>';
      $html .= '<li>';
      $html .= '<h2>' . __('Reading CSV files externally', 'wordtohtml-wp') . '</h2>';
      $html .= __('To read csv-file externally (remote from an url) you have to define at least one csv-file. Native wordpress http api is used as fetching remote files. It will fall back to CURL (if installed) if this approach does not work.', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<hr>';
      $html .= '<li>';
      $html .= __('<i>Example 1:</i> <strong>[wordtohtml_create source_files=”http://wibergsweb.se/sweden.docx”]</strong>', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<li>';
      $html .= __('This example would make an html table from the file sweden.docx if it exists on the root of wibergsweb.se - domain', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<hr>';

      $html .= '<li>';
      $html .= __('<i>Example 2:</i> <strong>[wordtohtml_create source_files=”http://wibergsweb.se/sweden.docx;http://wibergsweb.se/norway.docx”]</strong>', 'wordtohtml-wp');
      $html .= '</li>';    
      $html .= '<li>';
      $html .= __('This example would make an html table from the files sweden.docx and norway.docx if these files exists on the root of wibergsweb.se - domain', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<hr>';
      $html .= '</ul>';

    $html .= '<ul>';
      $html .= '<li>';
      $html .= '<h2>' . __('Debugging', 'wordtohtml-wp') . '</h2>';
      $html .= __('What differs Word to html from many other plugins is it\'s availibilty to debug and identify possible problems with your csv/shortcode. Just add debug_mode="yes" to your shortcode and you will get a lot of output. If something goes wrong without debugging it will be a "silent error", not showing error(s) for your visitor. If debug_mode is set to yes, it might be such a simple thing that the file does not exist.', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<hr>';
      $html .= '<li>';
      $html .= __('<i>Example 1:</i> <strong>[wordtohtml_create debug_mode="yes" source_files=”charity.docx”]</strong>', 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<li>';
      $html .= __('The plugin could give some clues what\'s going on and in some cases even some suggestions' , 'wordtohtml-wp');
      $html .= '</li>';
      $html .= '<hr>';

      $html .= '</ul>';
      
      $html .= '<h2>Support</h2>';
      $html .= '<p>';
      $html .= 'There are actual a lot more you could do with this plugin (styling, sorting columns, including, excluding data etc). To review the full list of attributes go to: ';
      $html .= '<a target="_blank" href="https://wordpress.org/plugins/word-to-html/#installation">Look into more options/attributes that is possible to do with this plugin.</a>';
      $html .= '</p>';
      
      $html .= '<p>If you need any help please don\'t hesitate to contact me (If you want to donate (the plugin is free but donations are welcome) contact me as well) , Gustav Wiberg at Wibergs Web. (<a href="mailto:info@wibergsweb.se">info@wibergsweb.se</a>).';
      
      $html .= '</div>';
      echo $html;
    }
    
    
    /**
     * Detects the end-of-line character of a string.
     * 
     * @param string $str The string to check.
     * @return string The detected eol. If no eol found, use default eol from object
     */    
    private function detect_eol( $str )
    {
        static $eols = array(
            "\0x000D000A", // [UNICODE] CR+LF: CR (U+000D) followed by LF (U+000A)
            "\0x000A",     // [UNICODE] LF: Line Feed, U+000A
            "\0x000B",     // [UNICODE] VT: Vertical Tab, U+000B
            "\0x000C",     // [UNICODE] FF: Form Feed, U+000C
            "\0x000D",     // [UNICODE] CR: Carriage Return, U+000D
            "\0x0085",     // [UNICODE] NEL: Next Line, U+0085
            "\0x2028",     // [UNICODE] LS: Line Separator, U+2028
            "\0x2029",     // [UNICODE] PS: Paragraph Separator, U+2029
            "\0x0D0A",     // [ASCII] CR+LF: Windows, TOPS-10, RT-11, CP/M, MP/M, DOS, Atari TOS, OS/2, Symbian OS, Palm OS
            "\0x0A0D",     // [ASCII] LF+CR: BBC Acorn, RISC OS spooled text output.
            "\0x0A",       // [ASCII] LF: Multics, Unix, Unix-like, BeOS, Amiga, RISC OS
            "\0x0D",       // [ASCII] CR: Commodore 8-bit, BBC Acorn, TRS-80, Apple II, Mac OS <=v9, OS-9
            "\0x1E",       // [ASCII] RS: QNX (pre-POSIX)
            "\0x15",       // [EBCDEIC] NEL: OS/390, OS/400
            "\r\n",
            "\r",
            "\n"
        );
        $cur_cnt = 0;
        $cur_eol = $this->default_eol;
        
        //Check if eols in array above exists in string
        foreach($eols as $eol){      
            $char_cnt = mb_substr_count($str, $eol);
                    
            if($char_cnt > $cur_cnt)
            {
                $cur_cnt = $char_cnt;
                $cur_eol = $eol;
            }
        }
        return $cur_eol;
    }

    
   
    /*
     *  convertarrayitem_encoding
     * 
     *  This function is used as a callback for walk_array and it changes
     *  characterencoding for each item in an array
     * 
     *  @param    array  $given_item           Arrayitem to translate encoding
     *  @return   N/A                          Change of arrayitem by reference
     *                 
     */  
    private function convertarrayitem_encoding( &$given_item ) 
    {       
        $encoding_to = $this->encoding_to;
        $encoding_from = $this->encoding_from;        
        
        $option_encoding = 0; //Only to encoding 
        if ( $encoding_from !== null && $encoding_to !== null ) 
        {
            $option_encoding = 1; //Both from and to encoding
        }
                         
        if ( $option_encoding === 1 )
        {
            if ( is_array($given_item) !== true ) 
            {
                $given_item = mb_convert_encoding($given_item, $encoding_to, $encoding_from);                  
            }
        }
        else if ( $option_encoding === 0 )
        {
            if ( is_array($given_item) !== true )
            {
                $given_item = mb_convert_encoding($given_item, $encoding_to);                       
            }
        }
                    
    }
    
    

    
            
    /*
     *   source_to_html
     * 
     *  This function creates html based on given source file(s)
     * 
     *  @param  string $attr             shortcode attributes
     *  @return   string                      html-content
     *                 
     */    
    public function source_to_html( $attrs ) 
    {
        $tables = array(); 
        
        $defaults = array(
            'html_id' > null,           //Is applied to article-tag (if skip_articletag is set to no)   
            'html_class' => null,       //Is applied to each section (one section per file)
            'path' => '', //This is the base path AFTER the upload path of Wordpress (eg. /2016/03 = /wp-content/uploads/2016/03)          
            'source_files' => null, //Files are be divided with sources_separator (file1;file2 etc). It's also possible to include urls to csv files. It's also possible to use a wildcard (example *.docx) for fetching all files from specified path. This only works when fetching files directly from own server.
            'eol_detection' => 'auto', //Use linefeed when using external files, Default auto = autodetect, CR/LF = Carriage return when using external files, CR = Carriage return, LF = Line feed
            'convert_encoding_from' => null, //If you want to convert character encoding from source. (use both from and to for best result) 
            'convert_encoding_to' => null, //If you want to convert character encoding from source. (use both from and to for best result)            
            'add_ext_auto' => 'yes', //If file is not included with .docx, then add .docx automatically if this value is yes. Otherwise, set no
            'skip_articletag' => 'yes', //If set to yes then <article>-tag will not be included (article is set as default in word)
            'debug_mode' => 'no'
        );

        //Extract values from shortcode and if not set use defaults above
        $args = wp_parse_args( $attrs, $defaults );
        extract ( $args );

        //Error handling
        if ( $debug_mode === 'yes') 
        {
            ini_set('xdebug.var_display_max_depth', '-1');
            ini_set('xdebug.var_display_max_children', '-1');
            ini_set('xdebug.var_display_max_data', '-1');
            
            var_dump ( $args );
            
            if ( $source_files === null) 
            {
                $this->errormessage = __('No source file(s) given. At least one file (or link) must be given', 'wordtohtml-wp');
                $this->error_notice();
                return;
            }   

            if (strlen( $path ) == 0 ) 
            {
                $this->errormessage = __('You must create a folder and use attribute path="folder on your webserver".<br>This path is also used when fetching external docx-documents.', 'wordtohtml-wp');
                $this->error_notice();
                return;
            }
                
            if ( strlen( $html_id ) > 0 && $skip_articletag == 'yes')
            {
                echo '<pre>';
                echo __('html id would not be set if skip_articletag is set to yes. <br>Set skip_articletag="no" if you wish that the article-tag with given id will be included in your html.','wordtohtml-wp');
                echo '</pre>';                            
            }
            
             
        }
        
        //Base upload path of uploads
        $upload_dir = wp_upload_dir();
        $upload_basedir = $upload_dir['basedir'];

        //If user has put some wildcard in source_files then create a list of files
        //based on that wildcard in the folder that is specified    
        if ( stristr( $source_files, '*' ) !== false ) 
        {            
            $files_path = glob( $upload_basedir . '/' . $path . '/'. $source_files);
            if ( $debug_mode === 'yes')
            {
                echo '<pre>';
                echo __('Files grabbed from wildcard: ' . $upload_basedir . '/' . $path . '/'. $source_files, 'wordtohtml-wp');
                echo '</pre>';
            }
            
            $source_files = '';
            foreach ($files_path as $filename) 
            {
                if ( $debug_mode === 'yes' )
                {
                    echo basename($filename) .  "<br>";
                }
                $source_files .= basename($filename) . ';';
            }
            if ( strlen($source_files) > 0) {
                $source_files = substr($source_files,0,-1); //Remove last semicolon
            }
            else 
            {
                if ( $debug_mode === 'yes')
                {
                    $this->errormessage = __('Wildcard set for source-files but no source file(s) could be find in specified path.', 'wordtohtml-wp');
                    $this->error_notice();   
                    return;
                }
            }
            
            if ( $debug_mode === 'yes')
            {
                echo '</pre>';
            }
        }

        //Find location of sources (if more then one source, user should divide them with 'sources_separator' (default semicolon) )
        //Example:  [stt_create path="2015/04" sources="bayern;badenwuertemberg"] 
        ///wp-content/uploads/2015/04/bayern.docx
        ///wp-content/uploads/2015/04/badenwuertemberg.docx        
        $sources = explode( ';', $source_files );

        //Create an array content from file(s)
        $content_arr = array();
        $style_arr = array();
        $nimages = array();
          
        foreach( $sources as $sourcekey => $s) 
        {
            //If $s(file) misses an extension add csv extension to filename(s)
            //if add extension auto is set to yes (yes is default)
            if (stristr($s, '.docx') === false && $add_ext_auto === 'yes') {
                $file = $s . '.docx';
            }
            else {
                $file = $s;
            }
          
            //Add array item with content from file(s)
        
            //If source file do not have http or https in it or if path is given, then it's a local file
            $local_file = true;
            
            if ( stristr($file, 'http') !== false || stristr($file, 'https') !== false )
            {
                $local_file = false;
            }                    
            
            //Load external file and add it into array
            if ( $local_file === false ) 
            {         
                $file_arr = false;
                               
                //Check if (external) file exists
                $wp_response = wp_remote_get($file);
                $ret_code = wp_remote_retrieve_response_code( $wp_response );
                $ret_message = wp_remote_retrieve_response_message( $wp_response );

                if ( $debug_mode === 'yes' )
                {                    
                    echo '<pre>'; 
                    echo __('Return code','wordtohtml-wp') . ': ' . $ret_code . '<br>';
                    echo __('Return message','wordtohtml-wp') . ': ' . $ret_message . '<br>';                        
                    echo '</pre>';
                }

                //200 OK               
                if ( $ret_code === 200)
                {
                    $external_source = $file;                    
                    $bfile_index = strrpos($file, '/');  //Find last occurence of / in string 
                    $basename_file = substr($file, $bfile_index+1); 
                    $file = $basename_file;
                    
                    $base_uploadpath = $upload_basedir . '/' . $path . '/';

                    if ( strlen( $path ) > 0 ) 
                    {
                        $file = $upload_basedir . '/' . $path . '/' . $file; //File from uploads folder and path
                    }
                    
                    //save external word-document to local storage
                    $result_copy = @copy( $external_source, $file);
                    
                    if ( $debug_mode == 'yes')
                    {
                        if ( $result_copy ) 
                        {
                            echo '<pre>';
                            echo __('copy success to ' . $file,'wordtohtml-wp');
                            echo '</pre>';
                        }
                        else 
                        {
                            echo '<pre>';
                            echo __('copy failed to ' . $file,'wordtohtml-wp');
                            echo __('Does the folder exists?','wordtohtml-wp');
                            
                            echo '</pre>';                            
                        }                            
                    }
                    
                    $local_file = true; //Use this localfile so we can docx (zip-document etc)
                    $file = $basename_file;                    
                }
                else
                {
                    if ( $debug_mode === 'yes') 
                    {
                        $this->errormessage = $file . ' not found';
                        $this->error_notice();
                    }
                }


            }
            
            //Load local file into content array
            if ( $local_file === true ) 
            {
                $basename_file = $file;                
                $base_uploadpath = $upload_basedir . '/' . $path . '/';
                
                if ( strlen( $path ) > 0 ) 
                {
                    $file = $upload_basedir . '/' . $path . '/' . $file; //File from uploads folder and path
                }
                else 
                {
                    $file = $upload_basedir . '/' . $file; //File directly from root upload folder
                }
                
                if (file_exists($file)) 
                {                                        
                    //Move entire file into an array
                    //A word-document is actually a zip-file where document.xml part of it contains the actual content
                    $result = file_get_contents( 'zip://' . $file . '#word/document.xml' );
                   
                    //$content_arr[] = simplexml_load_string($result,null, 0, 'w', true); 
                    
                    //Extract the word (docx) file so media-folder would be available
                    $zip = new ZipArchive;
                    $res = $zip->open( $file );
                    $word_resource = null;
                    $media_resource = null;
                    $new_images = array();
                    
                    if ($res === true) 
                    {
                        $word_resource = $upload_basedir . '/wordtohtml/'. $basename_file . '/';
                        $zip->extractTo( $word_resource );
                        $zip->close();
                        
                        
                        
                                
                        $result_relations = file_get_contents( $word_resource . 'word/_rels/' . 'document.xml.rels' );
                         
                        $media_resource = $word_resource . 'word/media';
  
                        if (file_exists($media_resource)) 
                        {
                            $m_images = array_diff( scandir( $media_resource ), array(".", "..") );
                            $mediaimages = array_values( $m_images );
                        }
                        else {
                            $mediaimages = array();
                        }
                        
                        foreach($mediaimages as $mi) 
                        {
                                if (substr($mi,-4,4) == 'jpeg' || substr($mi,-3,3) == 'jpg' || substr($mi,-3,3) == 'png' )
                                {
                                    $new_images[$sourcekey][] = $mi;
                                }
                                else {
                                    $new_images[$sourcekey][] = '';
                                }
                        }
                    }                    

                    $document = new DOMDocument();
                    $document->loadXML($result);
                    $xpath = new DOMXpath($document);
                    // register a prefix for the used namespace
                    $xpath->registerNamespace('wspace', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
                    

                    $document_rels = new DOMDocument();
                    $document_rels->loadXML($result_relations);
                    $xpath_rels = new DOMXpath($document_rels);
                    // register a prefix for the used namespace
                    $xpath_rels->registerNamespace('relspace', 'http://schemas.openxmlformats.org/package/2006/relationships');
                    

                    //Create an image array that reflects where in the DOM that images are
                    $pindex = 0;
                    $img_count = 0;
                    $uploaddir = wp_upload_dir();
                    $upload_to = $uploaddir['baseurl'] . '/wordtohtml/' . $basename_file . '/word/media';
                     
                    $html = '';
                    $createul_list = false;
                    $last_createullist = false;
                    $ul_level = 0;
                    $list_arr = array();
                    $ultype_arr = array();
                    foreach ($xpath->evaluate('//wspace:p|//wspace:tbl') as $index => $p_node) 
                    {
                        //Just do something when in node just under body-node.
                        if ( $p_node->parentNode->localName != 'body') {
                            continue;
                        }
                        
                        if (strlen($p_node->textContent)>0) 
                        {
                            switch($p_node->localName) {
                                case 'p':   
                                    $end_tag = '';
                                    $html .= '<p>';
                                    $child_nodes_p = $p_node->childNodes;
                                    $text = '';
                                    
                                   
                                    $createul_list = false;
                                    $list_start = false;
                                    foreach($child_nodes_p as $node_r) 
                                    {
                                        $ppr = false;
                                        $hyperlink = false;
 
                                        if ($node_r->localName == 'pPr') 
                                        {
                                            //$html .= '<strong>pPr::</strong>';
                                            $ppr = true;
                                        }
                                        
                                        $child_r = $node_r->childNodes;      
                                        
                                        
 
                                        $do_bold = false;
                                        $do_italic = false;
                                        $do_underline = false;
                                        if ($node_r->localName == 'r')
                                        {
                                            $tcncc = $node_r->childNodes;
                                            foreach($tcncc as $pprc) {
                                                if ( $pprc->localName == 'rPr')
                                                {
                                                  $rprchilds = $pprc->childNodes;
                                                    foreach($rprchilds as $rprc) 
                                                    {
                                                        if ( $rprc->localName == 'u') 
                                                        {
                                                            $do_underline = true;
                                                            $text .= '<span class="wthtml-underline">';
                                                        }   
                                                        if ( $rprc->localName == 'b') 
                                                        {
                                                            $do_bold = true;
                                                            $text .= '<strong>';
                                                        }
                                                        if ( $rprc->localName == 'i') 
                                                        {
                                                            $do_italic = true;
                                                            $text .= '<i>';
                                                        }                                                              
                                                    }
                                                
                                                }
                                                
                                            }
                                        }
                                        
                                        foreach($child_r as $rc) 
                                        {
                                            
                                            if ( $ppr === true && $rc->localName == 'numPr') {
                                                $levels = $rc->childNodes;
                                               
                                                foreach($levels as $l_list)
                                                {
                                                    if ($l_list->localName == 'ilvl') {
                                                        $ul_level = $l_list->attributes['val']->value;
                                                    }
                                                    
                                                    if ($l_list->localName == 'numId') {
                                                        $ul_type = $l_list->attributes['val']->value;
                                                        
                                                    }
                                                }
                                                
                                            }
                                            
                                            if ( $ppr === true && $rc->localName == 'pStyle') 
                                            {

                                               $name_pstyle = strtolower($rc->attributes->item(0)->value);
                                               
                                                if ( $name_pstyle == 'rubrik') {
                                                    $name_pstyle = 'rubrik1';
                                                }
                                                if ( $name_pstyle == 'heading') {
                                                    $name_pstyle = 'heading1';
                                                }                                                
                                                if ( substr($name_pstyle,0,7) == 'heading' || substr($name_pstyle,0,6) == 'rubrik' ) 
                                                {                                                
                                                    $heading_nr = substr($name_pstyle, -1,1);
                                                    if ( $heading_nr <= 6) 
                                                    {
                                                        $html .= '<h' . $heading_nr . '>';
                                                        $end_tag = '</h' . $heading_nr . '>';
                                                    }
                                                }
                                              
                                                
                                                if ( $name_pstyle == 'listparagraph' || $name_pstyle == 'liststycke') 
                                                {                                                    
                                                    $createul_list = true; 
                                                }
                                            }
                                            

                                            if ( $rc->localName == 'br')
                                            {
                                                $br_attributes = array();
                                                foreach($rc->attributes as $rca)
                                                {
                                                    if ( $rca->localName == 'type')
                                                    {
                                                        $br_attributes[] = $rca;    
                                                    }                                                    
                                                }
                                                
                                                $brvalue = null;
                                                if (isset($br_attributes[0])) 
                                                {
                                                    $brvalue = $br_attributes[0]->nodeValue;
                                                    if ( $brvalue == 'column' )
                                                    {
                                                        $text .= '</p><p class="column">';
                                                    }
                                                }
                                            }

                                            
                                            if ($rc->localName == 't') 
                                            {                                                     
                                                $text .= $rc->textContent;
                                            }
                                            
                                            if ( $rc->localName == 'br')
                                            {
                                                if ($brvalue === null) 
                                                {
                                                    $text .= '<br>';
                                                }
                                            }
                                            
                                            
                                            if ( $rc->localName == 'drawing' ) 
                                            {
                                                if ( strlen( $new_images[$sourcekey][$img_count])>0 ) 
                                                {
                                                    $html .= '<img src="' . $upload_to . '/' . $new_images[$sourcekey][$img_count] . '" alt="" title="">';
                                                }
                                                $img_count++;

                                            }   
                                            
                                            
                                        }
                                        
                                    if ( $do_bold === true) 
                                    {
                                        $text .= '</strong>';
                                    }                                        

                                    if ( $do_italic === true) 
                                    {
                                        $text .= '</i>';
                                    }  
                                    if ( $do_underline === true) 
                                    {
                                        $text .= '</span>';
                                    }                                      
                                        
                                        
                                 if ($node_r->localName == 'hyperlink') 
                                        {
                                            $hyperlink = true;
                                            if ( $hyperlink === true) 
                                            {
                                                $nattributes = array();
                                                $fattr = array();
                                                
                                                foreach($node_r->attributes as $na) 
                                                {                
                                                    $local_name = $na->localName;
                                                    $nattributes[$local_name] = $na->nodeValue;                                                    
                                                   
                                                }
                                                $idvalue_hyperlink = null;
                                                if (isset($nattributes['id'])) 
                                                {
                                                    $idvalue_hyperlink = $nattributes['id'];
                                                }

                                               
                                                
                                                $relations_arr = array();
                                                $rel_index = 0;
                                                foreach ($xpath_rels->evaluate('//relspace:Relationship') as $index_rel => $rel_node) 
                                                {
                                                    foreach($rel_node->attributes as $relkey=>$nvrel) 
                                                    {
                                                        $local_name = $nvrel->localName;
                                                        $relations_arr[$rel_index][$local_name] = $nvrel->nodeValue;
                                                    }
                                                
                                                    $rel_index++;
                                                }
                                                
                                                //Now we have the id of the hyperlink. This id is stored in _rels/word/document.xml.rels
                                                //where we get the actual url
                                                if ($idvalue_hyperlink !== null)
                                                {
                                                    foreach($relations_arr as $relations_item) 
                                                    {
                                                        if ($relations_item['Id'] == $idvalue_hyperlink) 
                                                        {
                                                            $target_url = $relations_item['Target'];
                                                            $target_mode = $relations_item['TargetMode'];

                                                            $url_target = '';
                                                            if ( $target_mode == 'External' ) 
                                                            {
                                                                $url_target = ' target="_blank"';
                                                            }
                                                            //var_dump($relations_item);
                                                            $text .= '<a href="' . $target_url . '"' . $url_target . '>';
                                                            $text .= $node_r->textContent;
                                                            $text .= '</a>';
                                                            break;
                                                        }
                                                    }
                                                }
                                                
                                                if (isset($nattributes['anchor'])) {
                                                    $text .= '<a href="#">' . $node_r->textContent . '</a>';
                                                   
                                                }
                                                
                                                
                                            }
                                        }                                        
                                    
                                    }
                                    
                                    //If ul is used a separate logic for that inserts into string $html
                                    if ( $createul_list === false)
                                    {
                                        $html .= $text;
                                    }
                              
                                    $html .= $end_tag;
                                    
                                    //Create listitem item in array
                                    if ($createul_list == true) 
                                    {
                                        $list_arr[][$ul_level] = $text;
                                        $ultype_arr[] = $ul_type;
                                    }
                                    
                                    //End of list
                                    if ( $createul_list === false && $last_createullist === true) {
                                        //$html .= '<pre>' . print_r($ultype_arr, true) . '</pre>';
                                        $ul_arr = array();
                                        $start_ul = 0;
                                        
                                        //Go through ultypes arr and increase level if change of type
                                        $last_ultype = $ultype_arr[0];
                                        foreach($ultype_arr as $ultype_key=>$ultype_item) 
                                        {
                                            if ( $ultype_item != $last_ultype )
                                            {
                                                //find highest level for this item with this key in listarr
                                                //and increase 1 (therefore count is like list_arr[$ultype_key][0])
                                                $highest_level = count( $list_arr[$ultype_key] );
                                                $level_down = $highest_level-1;
                                                
                                                //...and increase level for that key in listarr
                                                $list_arr[$ultype_key][$highest_level] = $list_arr[$ultype_key][$level_down];
                                                unset($list_arr[$ultype_key][$level_down]);                                                
                                            }
                                            else 
                                            {
                                                $last_ultype = $ultype_item;
                                            }
                                        }
                                        
                                        //$html .= '<pre>' . print_r($list_arr, true) . '</pre>';
                                        
                                        
                                        if ( $ultype_arr[0] == 1) {
                                            $ul_arr[] = '<ol>';
                                        }
                                        else {
                                            $ul_arr[] = '<ul>';
                                        }
                                        foreach($list_arr as $outer_listkey=>$outer_value) {
                                            foreach($outer_value as $level_list=>$inner_item) {  
                                                if ($outer_listkey>0 && $level_list == 0) {
                                                    if ( $start_ul >0) {
                                                        if ( $ul_type == 1) {
                                                            $ul_arr[] = '</ol>';
                                                        }
                                                        else {
                                                            $ul_arr[] = '</ul>';
                                                        }
                                                        $start_ul = 0;
                                                    }
                                                    
                                                    $ul_arr[] = '</li>';
                                                }
                                                
                                                $ul_type = $ultype_arr[$outer_listkey];                                                
                                                
                                                if ( $level_list == 0)
                                                {
                                                    $ul_arr[] = '<li>';
                                                    $ul_arr[] = $inner_item;   
                                                        
                                                }
                                                if ( $level_list >0)
                                                {  
                                                    if ($start_ul != $level_list) {
                                                        if ( $ul_type == 1) {
                                                            $ul_arr[] = '<ol>';
                                                        }
                                                        else {
                                                            $ul_arr[] = '<ul>';
                                                        }
                                                    }
                                                    
                                                    $ul_arr[] = '<li>' . $inner_item . '</li>';
                                                                                                        
                                                    $start_ul = $level_list;
                                                }
                                            }
                                        }
                                        $ul_arr[] = '</li>';
                                        if ( $ul_type == 1) {
                                            $ul_arr[] = '</ol>';
                                        }
                                        else {
                                            $ul_arr[] = '</ul>';
                                        }
                                        $start_ul = 0;

                                        //$html .= '<pre>' . print_r($ul_arr, true) . '</pre>';
                                        
                                        $ul_html = '';
                                        foreach($ul_arr as $li_litem) {
                                            $ul_html .= $li_litem;
                                        }
                                        $html .= $ul_html;
                                        $list_arr = array();
                                    }

                                    //for check start ul and end ul
                                    $last_createullist = $createul_list;
                                        
                                    
                                    $html .= '</p>';
                                    
                                    
                                    break;
                                    
                                case 'tbl':
                                    $html .= '<table><tbody>';
                                    $table_row = $p_node->childNodes;
                                    foreach($table_row as $tr) 
                                    {
                                        if ( $tr->localName == 'tr') {
                                            $html .= '<tr>';
                                            $table_col = $tr->childNodes;
                                            $found_text = false;
                                            
                                            $col_tag = 'td';
                                            $nodes_arr = array();
                                            $nodes_header = array();
                                            $do_header = false;
                                             
                                            
                                            $do_header = false;
                                            $do_bold = false;
                                            $do_italic = false;
                                            $do_underline = false;
                                            foreach($table_col as $tc) 
                                            {
                                                foreach($tc->childNodes as $tcn) 
                                                {
                                                    foreach($tcn->childNodes as $tcnc)
                                                    {
                                                        if ($tcnc->localName == 'pPr')
                                                        {
                                                            $tcncc = $tcnc->childNodes;
                                                            foreach($tcncc as $pprc) {
                                                                $rprchilds = $pprc->childNodes;
                                                                foreach($rprchilds as $rprc) {
                                                                    if ( $rprc->localName == 'u') {
                                                                        $do_underline = true;
                                                                    }                                                                    
                                                                    if ( $rprc->localName == 'b') {
                                                                        $do_bold = true;
                                                                    }
                                                                    if ( $rprc->localName == 'i') {
                                                                        $do_italic = true;
                                                                    }                                                                    
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                
                                                
                                                
                                                if ( $tc->localName == 'trPr' )
                                                {
                                                    
                                                    $do_header = true;
                                                }
                                                else 
                                                {
                                                    
                                                    if ( $do_header === true ) 
                                                    {
                                                        $col_tag = 'th';
                                                        $do_header = false;
                                                    }
                                                    else {
                                                        $col_tag = 'td';
                                                    }
                                                }

                                                
                                                $cnt_tc = $tc->childNodes;
                                                
                                                $add_tc = '';
                                                foreach($cnt_tc as $ctc) 
                                                {
                                                    $local_name = $ctc->localName;
                                                    $gridspan_value = null;
                                                    
                                                    
                                                    if ( $local_name == 'tcPr') 
                                                    {                                                        
                                                        
                                                        $childnodes_ctc = $ctc->childNodes;
                                                        foreach($childnodes_ctc as $cnp) 
                                                        {                                                            
                                                            if ($cnp->localName == 'gridSpan') 
                                                            {
                                                                foreach($cnp->attributes as $gridspan) {
                                                                    $gridspan_value = $gridspan->nodeValue;                                                                    
                                                                }
                                                            }
                                                        }
                                                        if ( $gridspan_value !== null) 
                                                        {
                                                            $add_tc = ' colspan="' . $gridspan_value . '"';
                                                            break;
                                                        }
                                                    }
                                                    
                                                    if ( $local_name == 'p') 
                                                    {
                                                        //If text not found, tell so we can create empty cells
                                                        
                                                        if (is_object($ctc->getElementsByTagName("t"))) {
                                                            $found_text = true;
                                                        }                                                                                                                
                                                    }
                                                    
                                                }
                                                                                                
                                                
                                                if ( $found_text === true ) 
                                                {
                                                    $tc_text = '';
                                                    
                                                    if (strlen( $tc->textContent ) == 0) {
                                                        $html .= '<td>&nbsp;</td>';
                                                    }
                                                    else 
                                                    {
                                                        if ( $do_underline === true )
                                                        {
                                                            $tc_text .= '<span class="wthtml-underline">';
                                                        }
                                                        if ( $do_bold === true )
                                                        {
                                                            $tc_text .= '<strong>';
                                                        }
                                                        if ( $do_italic === true )
                                                        {
                                                            $tc_text .= '<i>';
                                                        }
                                                        
                                                        $tc_text .= $tc->textContent;
                                                        
                                                        if ( $do_italic === true )
                                                        {
                                                            $tc_text .= '</i>';
                                                        }
                                                        if ( $do_bold === true )
                                                        {
                                                            $tc_text .= '</strong>';
                                                        }  
                                                        if ( $do_underline === true )
                                                        {
                                                            $tc_text .= '</span>';
                                                        }                                                        
                                                        
                                                       
                                                        $html .= '<' . $col_tag . $add_tc . '>' . $tc_text . '</td>';
                                                    }
                                                }
                                            }
                                            $html .= '</tr>';
                                        }
                                    }
                                    $html .= '</tbody></table>';
                                    
                                    break; //end p switch
    
                            }
                            
                            $pindex++;
                        }
                        
                        
                    }

                    
                    //Render all files from content_arr array
                    $content_arr[] = $html;                    
                    foreach ($xpath->evaluate('//wspace:p|//wspace:tbl') as $index => $node) 
                    {
                        if ($node->localName == 'tbl') 
                        {
                            $tables[$sourcekey][] = $index;                            
                        }
                    }
                }
                else if ( $debug_mode === 'yes' ) 
                {
                    $this->errormessage = $file . ' ' . __('not found','wordtohtml-wp');
                    $this->error_notice();
                }
            }
        }        
                
        if ( count ( $content_arr) === 0 && $debug_mode === 'yes') 
        {
            $this->errormessage = __('No files found','wordtohtml-wp');
            $this->error_notice();
            return;
        }
        
        $row_values = array();
        foreach ( $content_arr as $row ) {
            $get_row1 = str_replace( '<p></p>', '', $row );            
            $get_row2 = str_replace( '<span class="wthtml-underline"></span>','', $get_row1 );

            $row_values[]= $get_row2;                    
        }        
        
        //If encoding is specified, then encode entire array to specified characterset
        if ( $convert_encoding_from !== null || $convert_encoding_to !== null )
        {
            if ( $debug_mode === 'yes' )
            {
                $encoding_error = false;
                if ( $convert_encoding_from !== null && $convert_encoding_to === null)
                {
                    echo '<strong>' . __('You must tell what encoding to convert to', 'wordtohtml-wp') . '</strong><br>';   
                    $encoding_error = true;
                }
                
                if ( $convert_encoding_from !== null ) 
                {
                if (in_array($convert_encoding_from, mb_list_encodings() ) === false) 
                {                        
                    echo __('Convert FROM encoding is not supported', 'wordtohtml-wp') . '<br>';   
                    $encoding_error = true;
                }
                }

                if ( $convert_encoding_to !== null ) 
                {
                if (in_array($convert_encoding_to, mb_list_encodings() ) === false)
                {
                    echo __('Convert TO encoding is not supported', 'wordtohtml-wp') . '<br>';      
                    $encoding_error = true;
                }
                }
                
                if ( $encoding_error === true )
                {
                    echo __('Supported encodings:');
                    var_dump( mb_list_encodings() );                     
                }
                
            }
            
            $this->encoding_from = $convert_encoding_from;
            $this->encoding_to = $convert_encoding_to;        
            array_walk_recursive($row_values, array($this, 'convertarrayitem_encoding') );
        }
        

        if ( $debug_mode === 'yes') 
        {
            echo __('<h2>Showing row values</h2>', 'wordtohtml-wp');
            var_dump ( $row_values );            
        }

        //Create table
        if ( isset($html_id) ) 
        {
            $htmlid_set = 'id="' .  $html_id . '" '; 
        }
        else 
        {
            $htmlid_set = '';
        }
        
        if ( isset($html_class) ) 
        {
            $html_class = ' ' . $html_class;
        }
        else 
        {
            $html_class = '';
        }
        
        $html = '';
        if ( $skip_articletag == 'no' )
        {
            $html = '<article ' . $htmlid_set . '>';
        }        
        
        foreach ( $row_values as $source_index=>$html_item) 
        {                       
            $html .= '<section class="' . $html_class . ' wordtohtmlsource' . $source_index . '">';
            $html .= $html_item;
            $html .= '</section>';
            $source_index++;
        }  
        
        if ( $skip_articletag == 'no' )
        {        
            $html .= '</article>';
        }
       
        return $html;
    }

  }
        
$wordtohtmlwp = new wordtohtmlwp();
}