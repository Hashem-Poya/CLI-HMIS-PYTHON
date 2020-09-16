from openpyxl import Workbook, load_workbook
import time


class Excel_File_Handler:

	def __init__(self):
		
		self.customer_workbook = Workbook()
		self.room_workbook = Workbook()
		self.food_workbook = Workbook()

		self.customer_worksheet = self.customer_workbook.active
		self.room_sheet = self.room_workbook.active
		self.food_sheet = self.food_workbook.active

		self.customer_header_data = ['ID', 'Name', 'Address', 'Check-In-Date', 'Check-Out-Date', 'Room-ID', 'Food-ID']
		self.room_header_data = ['ID', 'Room-Number', 'Price', 'Is_Reserved']
		self.food_header_data = ['ID', 'Name', 'Price']

		self.customer_worksheet.append(self.customer_header_data)
		self.room_sheet.append(self.room_header_data)
		self.food_sheet.append(self.food_header_data)

		self.customer_workbook.save('customer_excel_file.xlsx')
		self.room_workbook.save('room_excel_file.xlsx')
		self.food_workbook.save('food_excel_file.xlsx')



	def load_customer_workbook(self):
		self.customer_book = load_workbook('customer_excel_file.xlsx')
		self.customer_sheet = self.customer_book.active	


	def load_room_workbook(self):
		self.room_book = load_workbook('room_excel_file.xlsx')
		self.room_sheet = self.room_book.active	
	

	def load_food_workbook(self):
		self.food_book = load_workbook('food_excel_file.xlsx')
		self.food_sheet = self.food_book.active	



	def customer_inputs(self):
		self.cust_id = input('Enter Customer ID: ')
		self.cust_name = input('Enter Customer Name: ')
		self.cust_address = input('Enter Customer Address: ')

			
	def room_inputs(self):
		self.room_id = input('Enter Room ID: ')
		self.room_number = input('Enter Room Number: ')
		self.room_price = input('Enter Room Price: ')

	def food_inputs(self):
		self.food_id = input('Enter Food ID: ')
		self.food_name = input('Enter Food Name: ')
		self.food_price = input('Enter Food Price: ')




	def add_customer(self):

		'''
			check if room exists, add customer and pick a room by ID
			check if food exists, pick a food by ID else add room and food
		'''

		self.load_customer_workbook()
		self.customer_inputs()

		self.r_id = ''
		self.f_id = ''

		if self.check_if_room_exists():

			print()
			self.list_rooms()
			print()
			self.r_id = input('Enter the Room ID for the Customer: ')

		else:
			print()
			print('Not Available Rooms')
			print()
			self.r_id = self.add_room()

		if self.check_if_food_exists():

			print()
			self.list_foods()
			print()
			
			self.f_id = input('Enter the Food ID for the Customer: ')

		else:

			self.f_id = self.add_food()


		self.customer_sheet.append(
			[
			self.cust_id, self.cust_name, self.cust_address,
			time.strftime('%Y-%m-%d %H:%m:%S') , '',
			self.r_id, self.f_id
			])

		self.customer_book.save('customer_excel_file.xlsx')
		self.reserve_room(int(self.r_id) + 1)



	def list_customers(self):
		self.load_customer_workbook()
		self.customer_sheet.dimensions
		self.print_data(self.customer_sheet.max_row, self.customer_worksheet.max_column, self.customer_sheet)


	# Done.
	def add_room(self):
		self.load_room_workbook()
		self.room_inputs()
		self.room_sheet.append([self.room_id, self.room_number, self.room_price, 'No'])
		self.room_book.save('room_excel_file.xlsx')
		return self.room_id

	# Done.
	def list_rooms(self):
		self.load_room_workbook()
		self.room_sheet.dimensions
		self.print_data(self.room_sheet.max_row, self.room_sheet.max_column, self.room_sheet)


	# For Reusing...
	def print_data(self, maximum_row, maximum_column, sheet):
		print()
		for r in range(1, maximum_row + 1):
			for c in range(1, maximum_column + 1):
				 print(sheet.cell(row=r, column=c).value, end = '	  ')
		print()


	# Done.
	def add_food(self):
		self.load_food_workbook()
		self.food_inputs()
		self.food_sheet.append([self.food_id, self.food_name, self.food_price])
		self.food_book.save('food_excel_file.xlsx')
		return self.food_id

	# Done.
	def list_foods(self):
		self.load_food_workbook()
		self.food_sheet.dimensions
		print('\nAvailable Food Lists:\n')
		self.print_data(self.food_sheet.max_row, self.food_sheet.max_column, self.food_sheet)
		print()

	def reserve_room(self, room_id):
		self.load_room_workbook()
		self.room_sheet.dimensions
		self.room_sheet['D{id}'.format(id=str(room_id))] = 'Yes'
		self.room_book.save('room_excel_file.xlsx')


	# Done. (IF Room Exists Return TRUE, Else Return FALSE)
	def check_if_room_exists(self):
		self.load_room_workbook()
		self.room_sheet.dimensions
		self.room_sheet.max_column

		self.room_exists = False

		for r_row in range(1, self.room_sheet.max_row + 1):
			for r_column in range(1, self.room_sheet.max_column + 1):
				if 'No' in self.room_sheet.cell(row=r_row, column=r_column).value:
					self.room_exists = True

		return self.room_exists


	def check_if_food_exists(self):
		self.load_food_workbook()
		self.food_sheet.dimensions
		if self.food_sheet.max_row > 1:
			print()
			print('Available foods')
			print()
			return True
		else:
			print()
			print('Not Available Foods')
			print()
			return False


