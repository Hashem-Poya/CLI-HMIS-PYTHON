from excel_file_handler import *
import sys

def main():

	file_handler = Excel_File_Handler()

	while True:

		print()
		print()
		print()
		print('1. Add Customer to System.')
		print('2. List all Customer Information.')
		print('3. Add Room to System.')
		print('4. List All Rooms.')
		print('5. Add Food to System.')
		print('6. List All Foods')
		print('7 Generate the Bill for Customer.')
		print('q to Exit.')
		print()
		print()
		print()

		operation = input('Enter the operation: ')

		if operation == '1':
			
			file_handler.add_customer()

		elif operation == '2':
			
			file_handler.list_customers()
		
		elif operation == '3':

			file_handler.add_room()
		
		elif operation == '4':
		
			file_handler.list_rooms()
		
		elif operation == '5':
		
			file_handler.add_food()

		elif operation == '6':

			file_handler.list_foods()

		elif operation == '7':

			# Calculate & Generate the bill
			file_handler.calculate_and_generate_bill()


		elif operation == 'q':
			sys.exit()

		else:
			print()
			print('Invalid Operation')
			print()

main()
