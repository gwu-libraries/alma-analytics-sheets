// Object that stores general variables as well as specific parameters for each Analytics report needed 
var configObj = {
	apikey: '',
	url: 'https://api-na.hosted.exlibrisgroup.com/almaws/v1/analytics/reports',
	spreadsheet: '',
	reports: [
		{
			sheet: 'orders_by_title',
			path: '/shared/The George Washington University/top_textbooks/orders_by_title',
			columns: ['Active Course Code',
						'PO Line Reference',
						'Title',
						'MMS Id',
						'Transaction Expenditure Amount',
						'PO Line Quantity',
						'List Price',
						'PO Line Total Price',
						'Invoice Line Quantity',
						'Invoice Line-Number'
					]
		},
		{
			sheet: 'titles_by_course',
			path: '/shared/The George Washington University/top_textbooks/titles_by_course',
			columns: ['Course Code',
					'Title',
					'MMS Id'
					]

		},
		{
			sheet: 'usage_by_item',
			path: '/shared/The George Washington University/top_textbooks/usage_by_item',
			columns: ['Barcode',
						'MMS Id',
						'Num of events'
					]
		},
		{
			sheet: 'items_by_title',
			path: '/shared/The George Washington University/top_textbooks/items_by_title',
			columns: ['Item Id', 
						'Holding Id', 
						'MMS Id', 
						'Title', 
						'Temporary Location Name', 
						'Barcode', 
						'num_items_by_title'
					]

		}
	]
}




