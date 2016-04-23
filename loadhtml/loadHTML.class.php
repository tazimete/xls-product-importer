<?php
/**
 * Fetch HTML from a give URL.
 *
 * @author   Md. Rafiqul Islam
 */
class loadHTML
{
	private $document;
	private $url;
	private $base;


	/**
	 * Creates a new image finder object.
	 */
	public function __construct($url)
	{
		// Store url
		$this->url = $url;
	}


	/**
	 * Loads the HTML from the url if not already done.
	 */
	public function load()
	{
		// Return if already loaded
		if($this->document)
				return;
		
		// Get the HTML document
		$this->document = self::get_document($this->url);

		// Get the base url
		$this->base = self::get_base($this->document);
		if( ! $this->base)
				$this->base = $this->url;
	}


	/**
	 * Returns an array with all the images found.
	 */
	public function get_images()
	{
		// Makes sure we're loaded
		$this->load();

		// Image collection array
		$images = array();
		
		// For all found img tags
		foreach($this->document->getElementsByTagName('img') as $img)
		{
				// Extract what we want
				$image = array
				(
						'src' => self::make_absolute($img->getAttribute('src'), $this->base),
				);
				
				// Skip images without src
				if( ! $image['src'])
						continue;

				// Add to collection. Use src as key to prevent duplicates.
				$images[$image['src']] = $image;
		}

		// Return values
		return array_values($images);
	}
	
	public function get_image_by_id($id)
	{
		// Makes sure we're loaded
		$this->load();

		// Image src
		$_image = '';
		
		// For all found img tags
		foreach($this->document->getElementsByTagName('img') as $img)
		{
				// Extract what we want
				$image = array
				(
						'src' => self::make_absolute($img->getAttribute('src'), $this->base),
						'id'  => $img->getAttribute('id')
				);
				
				// Skip images without src
				if( ! $image['src'] || $image['id']!=$id)
						continue;

				// Add to collection. Use src as key to prevent duplicates.
				$_image = $image['src'];
		}

		// Return values
		return $_image;
	}


	/**
	 * Gets the html of a url and loads it up in a DOMDocument.
	 */
	private static function get_document($url)
	{
		// Set up and execute a request for the HTML
		$request = curl_init();
		curl_setopt_array($request, array
		(
				CURLOPT_URL => $url,
				
				CURLOPT_RETURNTRANSFER => TRUE,
				CURLOPT_HEADER => FALSE,
				
				CURLOPT_SSL_VERIFYPEER => TRUE,
				CURLOPT_CAINFO => 'cacert.pem',

				CURLOPT_FOLLOWLOCATION => TRUE,
				CURLOPT_MAXREDIRS => 10,
		));
		$response = curl_exec($request);
		curl_close($request);

		// Create DOM document
		$document = new DOMDocument();

		// Load response into document, if we got any
		if($response)
		{
				libxml_use_internal_errors(true);
				$document->loadHTML($response);
				libxml_clear_errors();
		}

		return $document;
	}



	/**
	 * Tries to get the base tag href from the given document.
	 */
	private static function get_base(DOMDocument $document)
	{
		$tags = $document->getElementsByTagName('base');

		foreach($tags as $tag)
				return $tag->getAttribute('href');

		return NULL;
	}


	/**
	 * Makes sure a url is absolute.
	 */
	private static function make_absolute($url, $base) 
	{
		// Return base if no url
		if( ! $url) return $base;

		// Already absolute URL
		if(parse_url($url, PHP_URL_SCHEME) != '') return $url;
		
		// Only containing query or anchor
		if($url[0] == '#' || $url[0] == '?') return $base.$url;
		
		// Parse base URL and convert to local variables: $scheme, $host, $path
		extract(parse_url($base));

		// If no path, use /
		if( ! isset($path)) $path = '/';
 
		// Remove non-directory element from path
		$path = preg_replace('#/[^/]*$#', '', $path);
 
		// Destroy path if relative url points to root
		if($url[0] == '/') $path = '';
		
		// Dirty absolute URL
		$abs = "$host$path/$url";
 
		// Replace '//' or '/./' or '/foo/../' with '/'
		$re = array('#(/\.?/)#', '#/(?!\.\.)[^/]+/\.\./#');
		for($n = 1; $n > 0; $abs = preg_replace($re, '/', $abs, -1, $n)) {}
		
		// Absolute URL is ready!
		return $scheme.'://'.$abs;
	}
}
