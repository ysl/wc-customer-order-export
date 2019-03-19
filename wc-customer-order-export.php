<?php
/*
 * Plugin Name: WC Customer Order Export
 * Description: Woocommerce Customer Order Export
 * Author:      Brian Lin
 * Version:     0.1
 * Text Domain: wc-customer-order-export
 * Domain Path: /languages/
 * License:     GPL v2 or later
 */

defined( 'ABSPATH' ) or die();

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

add_action( 'plugins_loaded', array( 'WC_Customer_Order_Export', 'get_instance' ) );

class WC_Customer_Order_Export {

	protected static $instance = null;

	private $order = null;

	public static function get_instance() {
		// If the single instance hasn't been set, set it now.
		if ( self::$instance == null ) {
			self::$instance = new self;
		}

		return self::$instance;
	}

	public function __construct() {
		// Load plugin text domain
		add_action( 'init', array( $this, 'load_plugin_textdomain' ) );

		// Add order action in order list.
		// add_filter( 'woocommerce_admin_order_actions', array( $this, 'add_order_action' ), 99, 2 );
		add_action( 'woocommerce_admin_order_actions_end', array( $this, 'add_order_action2' ), 99 );
		add_action( 'admin_head', array( $this, 'add_order_action_button_css' ) );
		add_action( 'wp_ajax_customer_order_export', array( $this, 'customer_order_export' ) );

		// Add order action for individual order.
		add_action( 'woocommerce_order_actions', array( $this, 'custom_wc_order_action' ) );
		add_action( 'woocommerce_order_action_custom_action', array( $this, 'custom_action' ) );
	}

	public function load_plugin_textdomain() {
		$domain = 'wc-customer-order-export';
		$locale = apply_filters( 'plugin_locale', get_locale(), $domain );

		load_textdomain( $domain, trailingslashit( WP_LANG_DIR ) . $domain . '/' . $domain . '-' . $locale . '.mo' );
		load_plugin_textdomain( $domain, FALSE, basename( dirname( __FILE__ ) ) . '/languages' );
	}
	
	// Not work
	// public function add_order_action( $actions, $order ) {
	// 	$actions[] = [
	// 		'action' => 'export_order',
	// 		'url' => admin_url( 'admin-ajax.php?action=customer_order_export&order_id=' . $order->get_id() ),
	// 		'name' => __( 'Download to Xlsx', 'wc-customer-order-export' ),
	// 	];

	// 	return $actions;
	// }

	public function add_order_action2( $order ) {
		$action = 'export_order';
		$name = __( 'Download to Xlsx', 'wc-customer-order-export' );
		$url = admin_url( 'admin-ajax.php?action=customer_order_export&order_id=' . $order->get_id() );
		printf( '<a class="button tips view %1$s" href="%2$s" data-tip="%3$s">%4$s</a>', esc_attr( $action ), $url, esc_attr( $name ), esc_html( $name ) );
	}

	public function add_order_action_button_css() {
	    echo '<style>.column-wc_actions .export_order::after, .order_actions .export_order::after { font-family: woocommerce !important; content: "\e02e" !important; }</style>';
	}

	public function customer_order_export() {
		if ( ! isset( $_GET['order_id'] ) ) {
			exit;
		}

		$order = wc_get_order( $_GET['order_id'] );
		if ( ! $order ) {
			exit;
		}

		$this->order = $order;
		$this->download_order();
	}

	public function custom_wc_order_action( $actions ) {
		if ( is_array( $actions ) ) {
			$actions['custom_action'] = __( 'Download to Xlsx', 'wc-customer-order-export' );
		}

		return $actions;
	}

	public function custom_action( $order ) {
		$this->order = $order;
		$this->download_order();
	}

	private function download_order( ) {
		require_once dirname( __FILE__ ) . '/includes/spreadsheet/vendor/autoload.php';

		$filename = 'order-' . $this->order->get_id() . '.xlsx';

		$spreadsheet = new Spreadsheet();
		$this->compose_sheet( $spreadsheet );

		$writer = new Xlsx( $spreadsheet );

		// Redirect output to a client’s web browser (Xlsx)
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="' . $filename . '"');
		header('Cache-Control: max-age=0');
		// If you're serving to IE 9, then the following may be needed
		header('Cache-Control: max-age=1');

		// If you're serving to IE over SSL, then the following may be needed
		header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
		header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
		header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
		header('Pragma: public'); // HTTP/1.0

		$writer = IOFactory::createWriter( $spreadsheet, 'Xlsx' );
		$writer->save( 'php://output' );
		exit;
	}

	private function compose_sheet( $spreadsheet ) {
		$spreadsheet->setActiveSheetIndex( 0 );
		$active_sheet = $spreadsheet->getActiveSheet();

		// Page Setting
		$spreadsheet->getActiveSheet()->getPageSetup()->setOrientation( PageSetup::ORIENTATION_PORTRAIT );
		$spreadsheet->getActiveSheet()->getPageSetup()->setFitToWidth( 1 );
		$spreadsheet->getActiveSheet()->getPageSetup()->setFitToHeight( 1 );
		$spreadsheet->getActiveSheet()->getPageMargins()->setTop( 0.38 );
		$spreadsheet->getActiveSheet()->getPageMargins()->setRight( 0.38 );
		$spreadsheet->getActiveSheet()->getPageMargins()->setLeft( 0.38 );
		$spreadsheet->getActiveSheet()->getPageMargins()->setBottom( 0.38 );

		// Default setting.
		$active_sheet->getDefaultColumnDimension()->setWidth( 15 );
		$active_sheet->getColumnDimension( 'A' )->setWidth( 20 );
		$active_sheet->getColumnDimension( 'E' )->setWidth( 7 );

		$order = $this->order;

		// Border.
		$outline_border = [
			'borders' => [
				'outline' => [
					'borderStyle' => Border::BORDER_THIN,
					'color' => ['argb' => 'FF000000'],
				],
			],
		];
		$all_border = [
			'borders' => [
				'allBorders' => [
					'borderStyle' => Border::BORDER_THIN,
					'color' => ['argb' => 'FF000000'],
				],
			],
		];

		// Get billing name, address, phone.
		$name = $order->get_billing_first_name() . ' ' . $order->get_billing_company() . ' ' . $order->get_billing_address_1() . ' 老師收';
		$address = $order->get_billing_city();
		$phone = $order->get_billing_phone();
		$email = $order->get_billing_email();

		$active_sheet->setCellValue( 'A1', $name );
		$active_sheet->setCellValue( 'A2', $address );
		$active_sheet->setCellValue( 'A3', $phone );
		$active_sheet->setCellValue( 'A4', $email );
		$active_sheet->getStyle( 'A1:A4' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
		$active_sheet->getStyle( 'A1:A4' )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text
		$active_sheet->getStyle( 'A1:A4' )->getAlignment()->setWrapText( false );

		// Date
		$active_sheet->setCellValue( 'G2', '訂購日期' );
		$active_sheet->setCellValue( 'G3', $order->order_date );
		$active_sheet->getStyle( 'G2:G3' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( 'G2:G3' )->applyFromArray( $all_border );

		$active_sheet->setCellValue( 'A6', '奇米熊-出貨明細表' );
		$active_sheet->getStyle( 'A6' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( 'A6' )->getFont()->setSize( 18 );
		$active_sheet->getStyle( 'A6' )->getFont()->setBold( true );
		$active_sheet->mergeCells( 'A6:G6' );
		$active_sheet->getStyle( 'A6:G6' )->applyFromArray( $outline_border );
		
		// Get items
		$active_sheet->setCellValue( 'A8', '品名' );
		$active_sheet->setCellValue( 'B8', '數量' );
		$active_sheet->setCellValue( 'C8', '單價' );
		$active_sheet->setCellValue( 'D8', '金額' );
		$active_sheet->getStyle( "A8:D8" )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );

		$shown_products = [];
		$variable_products = [];
		foreach ( $order->get_items() as $item_id => $item_product ) {
			$product = $item_product->get_product();
			$total = $item_product->get_total();
			$is_gift = ( $total == 0 );
			if ( ! $is_gift ) {
				$shown_products[] = array(
					'name' => $item_product->get_data()['name'], //str_replace( '<br/>', "\n", $product->get_name() ),
					'quantity' => $item_product->get_quantity(),
					'total' => $total,
				);
			}

			// Check if variation product.
			if ( $product->is_type( 'variation' ) ) {  // Should we need original product?
				// Get the common data in an array:
				$item_product_data = $item_product->get_data();

				$variable_product = array(
					'name' => $item_product_data['name'],
					'quantity' => $item_product->get_quantity(),
					'attrs' => array(),
					'is_gift' => $is_gift, // Assume the gift is also a variable product.
				);

				// Get all metas
				foreach ( $item_product_data['meta_data'] as $wc_meta_data ) {
					$key = $wc_meta_data->key;
					$arr = null;
					// Integrate with product-input-fields-for-woocommerce plugin.
					if ( defined( 'ALG_WC_PIF_ID' ) && in_array( $key, ['_alg_wc_pif_global', '_alg_wc_pif_local'] ) ) {
						foreach ( $wc_meta_data->value as $field ) {
							// Check the type.
							if ( $field['type'] == 'file' ) {
								// Is image
								if ( isset( $field['_value']['type'] ) && strpos( $field['_value']['type'], 'image/' ) == 0 ) {
									$variable_product['attrs'][] = array(
										'name' => $field['title'],
										'type' => 'image',
										'value' => $field['_value']['_tmp_name']
									);
								} else {
									// No file or other type
									// Don't show.
								}
							} else {
								$variable_product['attrs'][] = array(
									'name' => $field['title'],
									'value' => $field['_value'],
								);
							}
						}
					} else {
						$variable_product['attrs'][] = array(
							'name' => urldecode( wc_attribute_label( $key ) ), // do urldecode for some attributes
							'value' => $wc_meta_data->value,
						);
					}
				}

				// Push
				$variable_products[] = $variable_product;
			}
		}

		// Sort by name.
		usort( $shown_products, function( $a, $b ) {
			return strcmp( $a['name'], $b['name'] );
		} );

		// Print out.
		$offset = 9;
		foreach ( $shown_products as $product ) {
			$active_sheet->setCellValue( "A{$offset}", str_replace( '<br/>', "\n", $product['name'] ) );
			$active_sheet->getStyle( "A{$offset}" )->getAlignment()->setWrapText( true );
			$active_sheet->setCellValue( "B{$offset}", $product['quantity'] );
			$active_sheet->setCellValue( "C{$offset}", (int)( $product['total'] / $product['quantity'] ) );
			$active_sheet->setCellValue( "D{$offset}", $product['total'] );

			$offset++;
		}

		// Subtotal
		// $active_sheet->setCellValue( "A{$offset}", '小計' );
		// $active_sheet->setCellValue( "D{$offset}", $order->get_subtotal() );
		// $offset++;
		$end_of_items_offset = $offset - 1;

		// Shipping
		$shipping_methods = $order->get_items( 'shipping' );
		$shipping_method = '';
		if ( count( $shipping_methods ) > 0 ) {
			$shipping_method = reset( $shipping_methods )->get_name();
		}
		$shipping_fee = $order->get_total_shipping();
		$active_sheet->setCellValue( "A{$offset}", "運費" );
		if ( $shipping_fee > 0 ) {
			$active_sheet->setCellValue( "B{$offset}", '1' );
		} else {
			$active_sheet->setCellValue( "B{$offset}", '0' );
		}
		$active_sheet->getStyle( "A{$offset}" )->getAlignment()->setWrapText( true );
		$active_sheet->setCellValue( "C{$offset}", $shipping_fee );
		$active_sheet->setCellValue( "D{$offset}", $shipping_fee );
		$offset++;

		// Total
		$active_sheet->setCellValue( "A{$offset}", '總計' );
		$active_sheet->setCellValue( "D{$offset}", $order->get_total() );
		$active_sheet->getStyle( "D{$offset}" )->getFont()->setSize( 18 );
		// Set border.
		$active_sheet->getStyle( "A8:D{$offset}" )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( "A8:D{$offset}" )->applyFromArray( $all_border );
		$offset++;

		// Set the itmes align to left.
		$active_sheet->getStyle( "A9:A{$end_of_items_offset}" )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );

		// Shipping method
		if ( $shipping_method ) {
			$active_sheet->setCellValue( "A{$offset}", "運送方式: {$shipping_method}" );
			$active_sheet->mergeCells( "A{$offset}:D{$offset}" );
			$active_sheet->getStyle( "A{$offset}:D{$offset}" )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
			$active_sheet->getStyle( "A{$offset}:D{$offset}" )->applyFromArray( $all_border );
			$offset++;
		}

		// Order ID
		$active_sheet->setCellValue( 'F8', '訂單編號' );
		$active_sheet->setCellValue( 'G8', '官網編號' );
		$active_sheet->setCellValue( 'G9', $order->get_id() );
		$active_sheet->getStyle( 'G9' )->getFont()->setSize( 18 );
		$active_sheet->mergeCells( 'F9:F10' );
		$active_sheet->mergeCells( 'G9:G10' );
		$active_sheet->getStyle( 'F8:G10' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( 'F8:G10' )->getAlignment()->setVertical( Alignment::VERTICAL_CENTER );
		// Set border.
		$active_sheet->getStyle( "F8:G10" )->applyFromArray( $all_border );
		$active_sheet->getStyle( "F8:G10" )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text

		// Payment
		$active_sheet->setCellValue( 'F12', $order->get_payment_method_title() );
		$active_sheet->mergeCells( 'F12:G13' );
		$active_sheet->getStyle( 'F12' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( 'F12' )->getAlignment()->setVertical( Alignment::VERTICAL_CENTER );
		$active_sheet->getStyle( 'F12:G13' )->getAlignment()->setWrapText( true );
		// Set border.
		$active_sheet->getStyle( "F12:G13" )->applyFromArray( $all_border );

		// Invoice
		$active_sheet->setCellValue( 'F15', '買方：統一編號' );
		$active_sheet->mergeCells( 'F15:G15' );
		$active_sheet->setCellValue( 'F16', $order->get_billing_last_name() );  // Put the 統一編號 in last name field.
		$active_sheet->getStyle( 'F16' )->getFont()->setSize( 18 );
		$active_sheet->mergeCells( 'F16:G17' );
		// Set border.
		$active_sheet->getStyle( 'F15:G17' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( "F15:G17" )->applyFromArray( $all_border );

		$offset += 2;

		// Variable products
		$gift_count = 0;
		// Sort by name.
		usort( $variable_products, function( $a, $b ) {
			return strcmp( $a['name'], $b['name'] );
		} );
		$image_id = 1;
		foreach ( $variable_products as $variable_product ) {
			if ( ! $variable_product['is_gift'] ) {
				$offset_start = $offset;
				$active_sheet->setCellValue( "A{$offset}", $variable_product['name'] );
				$offset++;

				$image_count = 0;
				$attr_count = count( $variable_product['attrs'] );
				foreach ( $variable_product['attrs'] as $attr ) {
					$active_sheet->setCellValue( "A{$offset}", $attr['name'] );
					if ( isset( $attr['type'] ) && $attr['type'] == 'image' ) {
						$active_sheet->setCellValue( "B{$offset}", '(如圖' . $image_id . ')' );
						$path = $attr['value'];
						if ( file_exists( $path ) ) {
							// Create new image
							$new_w = 200;
							$new_h = 100;
							$text_width = 60;
							$new_image = @imagecreatetruecolor( $new_w + $text_width, $new_h );
							$white = imagecolorallocate( $new_image, 255, 255, 255 );
							imagefill( $new_image, 0, 0, $white );

							// Add image id text.
							$text_color = imagecolorallocate( $new_image, 0, 0, 0 );
							$text = 'Image ' . $image_id;
							imagestring( $new_image, 2, 0, 0, $text, $text_color );

							// Evaluate new size.
							$size = getimagesize( $path );
							$src_w = $size[0];
							$src_h = $size[1];
							if ( ($src_w / $src_h) > ($new_w / $new_h) ) {
								$resize_w = $new_w;
								$resize_h = $src_h * $new_w / $src_w;
							} else {
								$resize_h = $new_h;
								$resize_w = $src_w * $new_h / $src_h;
							}

							// Copy image.
							$src_image = imagecreatefromstring( file_get_contents( $path ) );
							imagecopyresampled( $new_image, $src_image, $text_width, 0, 0, 0, $resize_w, $resize_h, $src_w, $src_h );

							// Paste to sheet.
							$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
							$drawing->setImageResource( $new_image );
							$drawing->setRenderingFunction( \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_PNG );
							$drawing->setMimeType( \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT );
							$drawing->setCoordinates( 'C' . $offset_start );
							$drawing->setWorksheet( $active_sheet );

							// Increase count.
							$image_id++;
							$image_count++;
						}
					} else {
						$active_sheet->setCellValue( "B{$offset}", $attr['value'] );
					}
					$offset++;
				}

				// Set border.
				$offset_end = $offset - 1;
				$active_sheet->getStyle( "A{$offset_start}:B{$offset_end}" )->applyFromArray( $all_border );
				$active_sheet->getStyle( "A{$offset_start}:B{$offset_end}" )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text

				$additional_offset = ($image_count * 5) - $attr_count;
				if ( $additional_offset > 0 ) {
					$offset += $additional_offset;
				}
				$offset++;

			} else {
				$gift_count++;
			}
		}

		$offset++;

		// Gift
		if ( $gift_count > 0 ) {
			$gift_groups = array();
			$active_sheet->setCellValue( "A{$offset}", '-- 以下為贈品 --' );
			$offset += 2;
			// Evaluate the count.
			foreach ( $variable_products as $variable_product ) {
				if ( $variable_product['is_gift'] ) {
					if ( ! isset( $gift_groups[$variable_product['name']] ) ) {
						$gift_groups[$variable_product['name']] = array();
					}
					// Loop for every attribute (should be only 1)
					foreach ( $variable_product['attrs'] as $attr ) {
						if ( ! isset( $gift_groups[$variable_product['name']][$attr['name']] ) ) {
							$gift_groups[$variable_product['name']][$attr['name']] = array();
						}
						// Evaluate the value count.
						if ( ! isset( $gift_groups[$variable_product['name']][$attr['name']][$attr['value']] ) ) {
							$gift_groups[$variable_product['name']][$attr['name']][$attr['value']] = $variable_product['quantity'];
						} else {
							$gift_groups[$variable_product['name']][$attr['name']][$attr['value']] += $variable_product['quantity'];
						}
					}
				}
			}

			// Print out
			foreach ( $gift_groups as $product_name => $gift_group ) {
				$active_sheet->setCellValue( "A{$offset}", $product_name );
				$active_sheet->getStyle( "A{$offset}:B{$offset}" )->applyFromArray( $all_border );
				$offset++;

				$offset_start = $offset;
				foreach ( $gift_group as $attr_name => $attr_val ) {
					$active_sheet->setCellValue( "A{$offset}", $attr_name );
					$active_sheet->setCellValue( "B{$offset}", '數量' );
					$offset++;

					// Sort by key.
					uksort( $attr_val, function( $a, $b ) {
						if ( is_numeric( $a ) && is_numeric( $b ) ) {
							return (int)$a > (int)$b;
						} else {
							return strcmp( $a, $b );
						}
					} );

					foreach ( $attr_val as $val => $count ) {
						$active_sheet->setCellValue( "A{$offset}", $val );
						$active_sheet->setCellValue( "B{$offset}", $count );
						$offset++;
					}
				}

				// Set border.
				if ( $offset > $offset_start) {
					$offset_end = $offset - 1;
					$active_sheet->getStyle( "A{$offset_start}:B{$offset_end}" )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
					$active_sheet->getStyle( "A{$offset_start}:B{$offset_end}" )->applyFromArray( $all_border );
					$active_sheet->getStyle( "A{$offset_start}:B{$offset_end}" )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text
				}
			}
		}

		$offset++;

		// Order note
		$active_sheet->setCellValue( "A{$offset}", '備註欄' );
		$active_sheet->mergeCells( "A{$offset}:G{$offset}" );
		$active_sheet->getStyle( "A{$offset}:G{$offset}" )->applyFromArray( $all_border );

		$offset++;

		$active_sheet->setCellValue( "A{$offset}", $order->get_customer_note() );
		$active_sheet->mergeCells( "A{$offset}:G{$offset}" );
		$active_sheet->getStyle( "A{$offset}:G{$offset}" )->applyFromArray( $all_border );
		$active_sheet->getStyle( "A{$offset}" )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text
	}

}
