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

		// Default setting.
		$active_sheet->getDefaultColumnDimension()->setWidth( 12 );
		
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
		$name = $order->get_billing_first_name() . $order->get_billing_last_name();
		$address = $order->get_billing_address_1() . $order->get_billing_address_2();
		$phone = $order->get_billing_phone();

		$active_sheet->setCellValue( 'A1', $name );
		$active_sheet->mergeCells( 'A1:C1' );
		$active_sheet->setCellValue( 'A2', $address );
		$active_sheet->mergeCells( 'A2:C2' );
		$active_sheet->setCellValue( 'A3', $phone );
		$active_sheet->mergeCells( 'A3:C3' );
		$active_sheet->getStyle( 'A1:D3' )->getAlignment()->setVertical( Alignment::HORIZONTAL_LEFT );
		$active_sheet->getStyle( 'A1:D3' )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text

		$active_sheet->setCellValue( 'A5', '出貨明細表' );
		$active_sheet->getStyle( 'A5' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( 'A5' )->getFont()->setSize( 16 );
		$active_sheet->mergeCells( 'A5:H5' );
		$active_sheet->getStyle( 'A5:H5' )->applyFromArray( $outline_border );
		
		// Get items
		$active_sheet->setCellValue( 'A7', '品名' );
		$active_sheet->mergeCells( 'A7:B7' );
		$active_sheet->setCellValue( 'C7', '數量' );
		$active_sheet->setCellValue( 'D7', '單價' );
		$active_sheet->setCellValue( 'E7', '金額' );

		$offset = 8;
		$variable_products = [];
		foreach ( $order->get_items() as $item_id => $item_product ) {
			//
			$product = $item_product->get_product();
			$product_name = $product->get_name();
			$quantity = $item_product->get_quantity();
			$total = $item_product->get_total();
			$is_gift = ( $total == 0 );
			if ( ! $is_gift ) {
				$row_num = $offset + $item_num;
				$active_sheet->setCellValue( "A{$row_num}", str_replace( '<br/>', "\n", $product_name ) );
				$active_sheet->getStyle( "A{$row_num}" )->getAlignment()->setWrapText( true );
				$active_sheet->mergeCells( "A{$row_num}:B{$row_num}" );
				$active_sheet->setCellValue( "C{$row_num}", $quantity );
				$active_sheet->setCellValue( "D{$row_num}", (int)( $total / $quantity ) );
				$active_sheet->setCellValue( "E{$row_num}", $total );

				$offset++;
			}

			// Check if variation product.
			if ( $product->is_type( 'variation' ) ) {
				// Get the common data in an array:
				$item_product_data = $item_product->get_data();

				$variable_product = array(
					'name' => $item_product_data['name'],
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
							$variable_product['attrs'][] = array(
								'name' => $field['title'],
								'value' => $field['_value'],
							);
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

		$active_sheet->mergeCells( "A{$offset}:E{$offset}" );
		$offset++;

		// Subtotal
		$active_sheet->setCellValue( "A{$offset}", '小計' );
		$active_sheet->mergeCells( "A{$offset}:B{$offset}" );
		$active_sheet->setCellValue( "E{$offset}", $order->get_subtotal() );
		$offset++;

		// Shipping
		$shipping_methods = $order->get_items( 'shipping' );
		$shipping_method = '';
		if ( count( $shipping_methods ) > 0 ) {
			$shipping_method = reset( $shipping_methods )->get_name();
		}
		$active_sheet->setCellValue( "A{$offset}", "運送方式: {$shipping_method}" );
		$active_sheet->mergeCells( "A{$offset}:B{$offset}" );
		$active_sheet->getStyle( "A{$offset}" )->getAlignment()->setWrapText( true );
		$shipping_fee = $order->get_total_shipping();
		$active_sheet->setCellValue( "E{$offset}", $shipping_fee );
		$offset++;

		// Total
		$active_sheet->setCellValue( "A{$offset}", '總計' );
		$active_sheet->mergeCells( "A{$offset}:B{$offset}" );
		$active_sheet->setCellValue( "E{$offset}", $order->get_total() );
		$active_sheet->getStyle( "E{$offset}" )->getFont()->setSize( 18 );

		// Set border.
		$active_sheet->getStyle( "A7:E{$offset}" )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( "A7:E{$offset}" )->applyFromArray( $all_border );

		$offset++;

		// Order ID
		$active_sheet->setCellValue( 'G7', '訂單編號' );
		$active_sheet->setCellValue( 'H7', '官網編號' );
		$active_sheet->setCellValue( 'G8', $order->get_id() );
		$active_sheet->mergeCells( 'G8:G9' );
		$active_sheet->mergeCells( 'H8:H9' );
		$active_sheet->getStyle( 'G7:H9' )->getAlignment()->setVertical( Alignment::VERTICAL_CENTER );
		// Set border.
		$active_sheet->getStyle( "G7:H9" )->applyFromArray( $all_border );
		$active_sheet->getStyle( "G7:H9" )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text

		// Payment
		$active_sheet->setCellValue( 'G11', $order->get_payment_method_title() );
		$active_sheet->getStyle( 'G11' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->getStyle( 'G11' )->getAlignment()->setVertical( Alignment::VERTICAL_CENTER );
		$active_sheet->mergeCells( 'G11:H12' );
		// Set border.
		$active_sheet->getStyle( "G11:H12" )->applyFromArray( $all_border );

		// Invoice
		$active_sheet->setCellValue( 'G14', '發票註記' );
		$active_sheet->mergeCells( 'G14:H14' );
		$active_sheet->getStyle( 'G14' )->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$active_sheet->mergeCells( 'G15:H16' );
		// Set border.
		$active_sheet->getStyle( "G14:H16" )->applyFromArray( $all_border );

		$offset += 2;

		// Variable products
		$gift_count = 0;
		foreach ( $variable_products as $variable_product ) {
			if ( ! $variable_product['is_gift'] ) {
				$offset_start = $offset;
				$active_sheet->setCellValue( "A{$offset}", $variable_product['name'] );
				$active_sheet->mergeCells( "A{$offset}:B{$offset}" );
				$offset++;
				foreach ( $variable_product['attrs'] as $attr ) {
					$active_sheet->setCellValue( "A{$offset}", $attr['name'] );
					$active_sheet->mergeCells( "A{$offset}:B{$offset}" );
					$active_sheet->setCellValue( "C{$offset}", $attr['value'] );
					$active_sheet->mergeCells( "C{$offset}:D{$offset}" );
					$offset++;
				}

				// Set border.
				$offset_end = $offset - 1;
				$active_sheet->getStyle( "A{$offset_start}:D{$offset_end}" )->applyFromArray( $all_border );
				$active_sheet->getStyle( "A{$offset_start}:D{$offset_end}" )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text

				$offset++;
			} else {
				$gift_count++;
			}
		}

		$offset++;

		// Gift
		if ( $gift_count > 0 ) {
			$active_sheet->setCellValue( "A{$offset}", '-- 以下為贈品 --' );
			$offset += 2;
			foreach ( $variable_products as $variable_product ) {
				if ( $variable_product['is_gift'] ) {
					$offset_start = $offset;
					$active_sheet->setCellValue( "A{$offset}", $variable_product['name'] );
					$active_sheet->mergeCells( "A{$offset}:B{$offset}" );
					$offset++;
					foreach ( $variable_product['attrs'] as $attr ) {
						$active_sheet->setCellValue( "A{$offset}", $attr['name'] );
						$active_sheet->mergeCells( "A{$offset}:B{$offset}" );
						$active_sheet->setCellValue( "C{$offset}", $attr['value'] );
						$active_sheet->mergeCells( "C{$offset}:D{$offset}" );
						$offset++;
					}

					// Set border.
					$offset_end = $offset - 1;
					$active_sheet->getStyle( "A{$offset_start}:D{$offset_end}" )->applyFromArray( $all_border );
					$active_sheet->getStyle( "A{$offset_start}:D{$offset_end}" )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_TEXT ); // Force text

					$offset++;
				}
			}
		}
	}

}
