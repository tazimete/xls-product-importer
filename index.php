
<?php 
	/*
	Plugin Name: xls-product-importer
	Plugin URI: http://wordpress.org/plugins/hello-dolly/
	Description: This is not just a plugin, it symbolizes the hope and enthusiasm of an entire generation summed up in two words sung most famously by Louis Armstrong: Hello, Dolly. When activated you will randomly see a lyric from <cite>Hello, Dolly</cite> in the upper right of your admin screen on every page.
	Author: NeshmediaBD
	Version: 1.0
	Author URI: http://ma.tt/
	*/

	require_once('phpExcel/PHPExcel.php');
	require_once('loadhtml/loadHTML.class.php');
	//define('XLSPI_PLUGINS_FILE_PATHS', plugins_url()."/xls-product-importer/index.php");
	
	ini_set('max_execution_time', -1);
	ini_set('memory_limit', -1);
	
	//Registering Arabic Tutor menu in Admin Page 
	add_action( 'admin_menu', 'register_xlspi_menu' );

	function register_xlspi_menu() {
		add_menu_page( 'xls-product-importer', ' XLS Product Importer', 'manage_options', 'xls_product_importer','show_xlspi_main_page' ); 	
	}
	
	//Add jQuery 
	function xlspi_enqueue_jquery(){
		wp_enqueue_script('jquery');
	}
	//add_action('admin_enqueue_scripts','xlspi_enqueue_jquery');

	//Showing content for mainmenu 
	function show_xlspi_main_page(){	?>
		<div class="wrap">
			<div class="xlspi-container">
				<h1>XLS Product Importer</h1>
				<br>
				<form id="form-category-frame" action="#" method="post" enctype="multipart/form-data">
				<table>
					<tr>
						<td><label for="xlspi-selected-post">Select product type : </label></td>
						<td><select class="input-select-large" name="xlspi-selected-post" id="xlspi-selected-post" class="xlspi-selected-post"> 
							<option value=" " selected="selected" onclick="showDropdown()">  </option>
							<option value="mirror" onclick="showDropdown()">Mirror</option>
							<option value="frame"  onclick="showDropdown()">Frame</option>
							</select> <!--<a onclick="showDropdown()" class="button button-primary button-large" id="xlspi-btn-select">Select</a>-->
						</td>
					</tr>
				</table>
				
				<div class="xlspi-form-container">
					<div class="xlspi-post-category-container">
						<div class="xlspi-post-category-frame" id="xlspi-post-category-frame">
							<table>
								<tr>
									<td> <label for="xlspi-category">Select Frame Category : </label></td>
									<td> &nbsp </td>
									<td>  <?php wp_dropdown_categories( 'taxonomy=mouldings_cat&hierarchical=1&name=xlspi-category-frame&class=input-select-large' ); ?></td>
								</tr>
							</table>
						</div>
						<div class="xlspi-post-category-mirror" id="xlspi-post-category-mirror">
							<table>
								<tr>
									<td> <label for="xlspi-category">Select Mirror Category : </label></td>
									
									<td>  <?php wp_dropdown_categories( 'taxonomy=mirror_cat&hierarchical=1&name=xlspi-category-mirror&class=input-select-large' ); ?></td>
								</tr>
							</table>
						</div>
						<div class="xlspi-file-uploader-container" id="xlspi-file-uploader-container">
							<table> 
								<tr>
									<td> <label for="xlspi-file-uploader">Upload XLS File : </label></td> 
									
									<td> <input type="file" name="xlspi-file-uploader"/> </td>
								</tr>
								<tr>
									 
									<td> <input type="submit" name="xlspi-file-uploader-submit" class="button button-primary button-large" value="Import"/> </td>
									<td> &nbsp </td> 
								</tr>
							</table>
						</div>
						<div class="xlspi-post-category-frame">
							
						</div>
					</div>
				</div>
				</form>
			</div>
		</div>
	<?php }
	
	
	//Receiving Excel FIle
	function receive_file_and_create_post(){
		if(isset($_POST['xlspi-file-uploader-submit']) && isset($_FILES)){
			//echo "File Received";
			$post_type = $_POST['xlspi-selected-post'];
			$post_category = $_POST['xlspi-category-frame'];
			$post_taxonomy_name = ($post_type == 'mirror' ? 'mirror_cat' : 'mouldings_cat');
			
			//echo $post_category; exit;
			
			$inputFileName = $_FILES['xlspi-file-uploader']['tmp_name'];
			//$inputFileType = $_FILES['xlspi-file-uploader']['type'];
			
			//print_r($_FILES); exit;
			
			if(is_uploaded_file($_FILES['xlspi-file-uploader']['tmp_name'])){
				//echo "File Upl;oaded Successfullyfdbgfdgbfdgb bfd gdf gdfg";
			}else{	
				print "Failed to upload file.................................................";
			}
				
			  try {
				$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
				$objReader = PHPExcel_IOFactory::createReader($inputFileType);
				$objPHPExcel = $objReader->load($inputFileName);
			} catch(Exception $e) {
				die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
			} 
			
			//  Get worksheet dimensions
			$sheet = $objPHPExcel->getSheet(0);
			$highestRow = $sheet->getHighestRow();
			$highestColumn = $sheet->getHighestColumn();
						
			//  Loop through each row of the worksheet in turn
			for ($row = 2; $row <= $highestRow; $row++) {
				//  Read a row of data into an array
				$rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, 
				NULL, TRUE, FALSE);
				
				$product_title = trim($rowData[0][0]);
				$product_code = trim($rowData[0][1]);
				$product_quantity = trim($rowData[0][2]);
				$product_fetaures = trim($rowData[0][3]);
				$product_specifications = trim($rowData[0][4]);
				$product_url = trim($rowData[0][5]);
				
				$product_specifications = explode("\n",$product_specifications);
				
				$_height = (isset($product_specifications[0]) ? explode(':',$product_specifications[0]) : '');
				$_height = (isset($_height[1]) ? (float)$_height[1] : '');
				$product_height = $_height;
				
				$_width = (isset($product_specifications[1]) ? explode(':',$product_specifications[1]) : '');
				$_width = (isset($_width[1]) ? (float)$_width[1] : '');
				$product_width = $_width;
				
				$_product_rebate_height = (isset($product_specifications[2]) ? explode(':',$product_specifications[2]) : "");
				$_product_rebate_height = (isset($_product_rebate_height[1]) ? (float)$_product_rebate_height[1] : '');
				$product_rebate_height = $_product_rebate_height;
				
				$_product_rebate_width = (isset($product_specifications[3]) ? explode(':',$product_specifications[3]) : "");
				$_product_rebate_width = (isset($_product_rebate_width[1]) ? (float)$_product_rebate_width[1] : '');
				$product_rebate_width = $_product_rebate_width;
				
				$_product_style = (isset($product_specifications[4]) ? explode(':',$product_specifications[4]) : "");
				$_product_style = (isset($_product_style[1]) ? trim($_product_style[1]) : '');
				$product_style = $_product_style;
				
				$_product_color = (isset($product_specifications[5]) ? explode(':',$product_specifications[5]) : "");
				$_product_color = (isset($_product_color[1]) ? trim($_product_color[1]) : '');
				$product_color = $_product_color;
				
				//echo $product_style;
				//echo $product_color; 

				//meta key
				$meta_key_product_code = ($post_type == 'mirror' ? 'mirror-product-code' : 'mouldings-product-code');
				$meta_key_product_height = ($post_type == 'mirror' ? 'mirror-height' : 'mouldings-height');
				$meta_key_product_width = ($post_type == 'mirror' ? 'mirror-width' : 'mouldings-width');
				$meta_key_product_rebate_height = ($post_type == 'mirror' ? 'mirror-rebate-height' : 'rebate-height');
				$meta_key_product_rebate_width = ($post_type == 'mirror' ? 'mirror-rebate-width' : 'rebate-width');
				$meta_key_product_style = ($post_type == 'mirror' ? 'mirror-style' : 'mouldings-style');
				$meta_key_product_color = ($post_type == 'mirror' ? 'mirror-colour' : 'mouldings-colour');
				$meta_key_product_fetaures = ($post_type == 'mirror' ? 'mirror-product-features' : 'mouldings-product-features');
				//$meta_key_product_price = ($post_type == 'mirror' ? 'mirror-price' : 'mouldings-price');
				
				/* echo '<div><pre>';
				//print_r($product_specifications);
				print_r($product_height);
				print_r($product_width);
				print_r($product_rebate_height);
				print_r($product_rebate_width);
				print_r($product_style);
				print_r($product_color);
				print_r($product_title);
				//print_r($rowData);
				echo '</pre></div>'; */
				
					//echo "Row: ".$row."- Col: ".($k+1)." = ".$v."<br />";
					//global $user_ID;
					global $wpdb;
					$check_meta = $wpdb->get_var( $wpdb->prepare("SELECT post_id FROM $wpdb->postmeta WHERE meta_key = %s AND meta_value = %s LIMIT 1" , $meta_key_product_code, $product_code ) );
					
					//print_r($check_meta->post_id); exit;
					
					if($check_meta){
						//Update post
						$updating_post = array(
							'ID'            => $check_meta,
							'post_title'    => $product_title,
							'post_content'  => ' ',
							'post_status'   => 'publish',
							'post_date'     => date( 'Y-m-d H:i:s' ),
							'post_author'   => 1,
							'post_type'     => $post_type
							//'category' => $post_category
						);
						$last_post_id = wp_update_post( $updating_post );
						if($last_post_id){
							//Adding meta value
							update_post_meta($last_post_id, $meta_key_product_code, $product_code);
							update_post_meta($last_post_id, $meta_key_product_height, $product_height);
							update_post_meta($last_post_id, $meta_key_product_width, $product_width);
							update_post_meta($last_post_id, $meta_key_product_rebate_height, $product_rebate_height);
							update_post_meta($last_post_id, $meta_key_product_rebate_width, $product_rebate_width);
							update_post_meta($last_post_id, $meta_key_product_style, $product_style);
							update_post_meta($last_post_id, $meta_key_product_color, $product_color);
							update_post_meta($last_post_id, $meta_key_product_fetaures, $product_fetaures);
							
							//$post_category_id = $post_category;
							wp_set_object_terms( $last_post_id, (int)$post_category, $post_taxonomy_name );
							
							//Adding attachment
							generate_attachment_url($product_url, $last_post_id);
						}
					
					}else{
						$inserting_post = array(
						'post_title'    => $product_title,
						'post_content'  => ' ',
						'post_status'   => 'publish',
						'post_date'     => date( 'Y-m-d H:i:s' ),
						'post_author'   => 1,
						'post_type'     => $post_type
						//'category' => $post_category
					);
					$last_post_id = wp_insert_post( $inserting_post ); 
					
					if($last_post_id){
						//Adding meta value
						add_post_meta($last_post_id, $meta_key_product_code, $product_code);
						add_post_meta($last_post_id, $meta_key_product_height, $product_height);
						add_post_meta($last_post_id, $meta_key_product_width, $product_width);
						add_post_meta($last_post_id, $meta_key_product_rebate_height, $product_rebate_height);
						add_post_meta($last_post_id, $meta_key_product_rebate_width, $product_rebate_width);
						add_post_meta($last_post_id, $meta_key_product_style, $product_style);
						add_post_meta($last_post_id, $meta_key_product_color, $product_color);
						add_post_meta($last_post_id, $meta_key_product_fetaures, $product_fetaures);
						
						//$post_category_id = $post_category;
						wp_set_object_terms( $last_post_id, (int)$post_category, $post_taxonomy_name );
						//Adding attachment
						generate_attachment_url($product_url, $last_post_id);
						
					}
				}
			} 
		}
	}
		
		add_action('init', 'receive_file_and_create_post');
		

	function generate_attachment_url($url, $post_id){
		//$url = 'http://www.antons.com.au/en/moulding/shop-by-style/all/adeline-goldblack-a84001-3';
		$html = new loadHTML($url);
		$image_url = $html->get_image_by_id('oucProductPictureViewer_imgProductFrame');
		$upload_dir = wp_upload_dir();
		//$file_url = $upload_dir;
		$filename = 'product_image-'.$post_id.'.jpg';

		if(wp_mkdir_p($upload_dir['path']))
			$file = $upload_dir['path'] . '/' . $filename;
		else
			$file = $upload_dir['basedir'] . '/' . $filename;

		/* if($image)
		echo '<img src="'.$image.'" />';
		else
		echo 'No image found'; */

		$fp = fopen ($file, 'w+');
		$ch = curl_init($image_url);
		// curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false); // enable if you want
		curl_setopt($ch, CURLOPT_FILE, $fp);          // output to file
		curl_setopt($ch, CURLOPT_FOLLOWLOCATION, 1);
		curl_setopt($ch, CURLOPT_TIMEOUT, 1000);      // some large value to allow curl to run for a long time
		curl_setopt($ch, CURLOPT_USERAGENT, 'Mozilla/5.0');
		// curl_setopt($ch, CURLOPT_VERBOSE, true);   // Enable this line to see debug prints
		$content = curl_exec($ch);

		curl_close($ch);                              // closing curl handle
		fclose($fp);
		
		//Adding attachment to post
		$wp_filetype = wp_check_filetype($filename, null );
		$attachment = array(
			'post_mime_type' => $wp_filetype['type'],
			'post_title' => sanitize_file_name($filename),
			'post_content' => '',
			'post_status' => 'publish'
		);
		$attach_id = wp_insert_attachment( $attachment, $file, $post_id );
		require_once(ABSPATH . 'wp-admin/includes/image.php');
		$attach_data = wp_generate_attachment_metadata( $attach_id, $file );
		wp_update_attachment_metadata( $attach_id, $attach_data );

		set_post_thumbnail( $post_id, $attach_id );
							
	}


?>

<style type="text/css">
.xlspi-container{
	width: 50%;
	margin:0px auto;
}

#xlspi-btn-select{
	text-align: center;
}

#xlspi-btn-select{
	width: 100px;
}

.input-select-large{
	width: 200px;
	text-align: left;
}

.xlspi-post-category-frame, .xlspi-post-category-mirror{
	margin-top: 40px;
	display: none;
}

#xlspi-file-uploader-container{
	display: none;
}

.xlspi-post-category-frame ul li ul{
	margin-left: 50px;
}

#option-parent{
	padding-left: 5px;
	width: 100%;
}

#option-children{
	padding-left: 20px;
}
</style>

<script type="text/javascript">
	
	//jQuery.noConflict();
	/* jQuery(document).ready(function($){
		$('#xlspi-btn-select').click(function(){
			$('.xlspi-post-category-mirror, .xlspi-post-category-frame').hide();
			var post = $('#xlspi-selected-post :selected').val();
			if(post){
				console.log(post);
			}else{
				console.log("Empty Value............");
			}
			
			if(post == 'mirror'){
				$('.xlspi-post-category-mirror').show();
			}else if(post == 'frame'){
				$('.xlspi-post-category-frame').show();
			}
		});
	}); */
	
	//Javscript
	//document.getElementById("xlspi-btn-select").onclick = showDropdown;
	
	function showDropdown(){
		var post = document.getElementById("xlspi-selected-post").value;
		var mirrorDiv = document.getElementById("xlspi-post-category-mirror");
		var frameDiv = document.getElementById("xlspi-post-category-frame");
		var fileUploaderContainerDiv = document.getElementById("xlspi-file-uploader-container");
		
		mirrorDiv.style.display = 'none';
		frameDiv.style.display = 'none';
		fileUploaderContainerDiv.style.display = 'none';
		
		if(post == 'mirror'){
			mirrorDiv.style.display = 'block';
			fileUploaderContainerDiv.style.display = 'block';
		}else if(post == 'frame'){
			frameDiv.style.display = 'block';
			fileUploaderContainerDiv.style.display = 'block';
		}
	}
	
</script>