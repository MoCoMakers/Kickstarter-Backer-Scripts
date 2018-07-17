
#Python3

import csv
import glob
import sys
import pycountry

Shipping_Data_Folder = "All Rewards - Jul 1 11pm - usps"
Shipping_Details_From_Kickstarter_Style_Export_folder="KickstarterStyleReport Jul 1"
LIST_FOCUS = "INTERNATIONAL"   	#Options: DOMESTIC, INTERNATIONAL, ALL
							#DOMESTIC will NOT show non-subscribed people
							#INTERNATIONAL will show blank values for people with no backer report
														
Main_Folder=r"D:\Programs\Hubic\Maker\Business\FPGA\Kickstarter Delivery\Rewards Reports"
MAKE_STAMPS_ORDER_FILES = True # If true please enter the below values:

#STAMPS.COM Order Variables
ORDER_DATE = "5/21/2018"
REQUESTED_SERVICE = "Package/Thick Envelope"
ITEM_WEIGHT_OZ = "7.7" #1.3
ITEM_LENGTH = "9.5"  #9.5
ITEM_WIDTH = "6.75"
ITEM_HEIGHT = "0.75"
PER_ITEM_COST = 20.0
#END STAMPS.COM Order Variables

#One PreSoldered Fipsy - Order Date 5/21/2018 - 1.3 oz - 9.5L 6.75W 0.75H - Cost - 20
#5LBS - Test Data
#6LBS - One Off Exceptions

files = glob.glob(Main_Folder+'\\'+Shipping_Data_Folder+'\\Origionals\\*.csv')
shipping_details_files = glob.glob(Main_Folder+'\\'+Shipping_Details_From_Kickstarter_Style_Export_folder+'\\*.csv')

def openCSV(file):
	csvfile = open(file, 'rt',encoding="utf8")
	csv_reader = csv.reader(csvfile, delimiter=',', quotechar='"')
	return csv_reader
	First=row[0]
	Last=row[2]
	Address1=row[4]
	Address2=row[5]
	Address3=row[6]
	City=row[7]
	State=row[8]
	Zip=row[9]
	Country_code=row[10]
	Quantity=row[17]
	Description=row[18]
	if Country_code:
		Country_name= pycountry.countries.get(alpha_2=Country_code).name
	else:
		Country_name=''
	Phone=row[12]
	Email=row[14]
	ReferenceNumber=row[15]	
	
	output_str=First+" "+Last+"\n"+Address1
	
	if Address2:
		output_str=output_str+"\n"+Address2
		
	if Address3:
		output_str=output_str+"\n"+Address3
	output_str=output_str+"\n"+City+", "+State+" "+Zip
	
	if not Address2:
		output_str=output_str+"\n"
	
	if not Address3:
		output_str=output_str+"\n"
	
	return output_str
	
if __name__=="__main__":
	if (sys.version_info > (3, 0)):
	 pass
	else:
	 raise RuntimeError('Please use Python 3 to run this script')

	shipping_details_dict = {}
	for shipping_details_csv in shipping_details_files:
		if "No reward" in shipping_details_csv:
			pass
		else:
			shipping_csv_reader=openCSV(shipping_details_csv)
			is_first=True
			
			for row in shipping_csv_reader:
				if is_first:
					is_first=False
				else:
					print(str(row))
					backer_number = row[0]
					shipping_details = row[24]
					print("Backer Number: "+str(backer_number))
					print("Description: "+str(shipping_details))
					shipping_details_dict[backer_number]=shipping_details
	
	for file in files:
		if "No reward - " in file:
			pass
		else:
			csv_reader=openCSV(file)
					
			is_first=True
			
			customers=[]
			
			for row in csv_reader:
				if is_first:
					is_first=False
				else:
					Country_code=row[10]
					if Country_code:
						Country_name= pycountry.countries.get(alpha_2=Country_code).name
					else:
						Country_name=''
						
					if Country_code == 'GB':
						Country_name = 'Great Britain'
					elif Country_code =='CZ':
						Country_name = 'Czech Republic'
					row[10]=Country_name
					
					Phone=row[12]
					print("Phone before: "+str(Phone))
					if len(Phone)>14:
						Phone = Phone.replace(" ","")
						Phone = Phone.replace("(","")
						Phone = Phone.replace(")","")
						Phone = Phone.replace("+","")
						Phone = Phone.replace(u"\u202C","") # Remove invisible character U+202C
						Phone = Phone.replace(u"\u202D","") # Remove invisible character U+202D
						Phone = Phone.replace("cell:","")
						Phone = Phone.replace("Cell:","")
						Phone = Phone.replace("-","")
						if Phone[:2]=="00":
							Phone = Phone[2:]
						
					row[12] = Phone
					print("Phone after: "+str(Phone))
					
					if u"\u202D" in Phone:
						raise Exception("Found hidden Character")
										
					print(str(row))
					
					Description = file.split('\\')[-1].split(' - ')[0].split('USD')[1].strip()
					if Description=="One Fipsy FPGA" or Description=="One Pre-soldered Fipsy FPGA":
						quantity=1
					elif Description=="Two Fipsy Pack" or Description=="Pre-soldered Fipsy FPGA Two-Pack":
						quantity=2
					elif Description=="Fipsy FPGA Four-Pack":
						quantity=4
					elif Description=="Fipsy FPGA Ten-Pack" or Description=="Pre-soldered FPGA Ten-Pack":
						quantity=10
					else:
						raise Exception("Failed to match quantity to file name")
					
					quantity=str(quantity)
					
					print(str(file))
					if LIST_FOCUS=="DOMESTIC" and Country_code!='US':
						pass
					elif LIST_FOCUS=="INTERNATIONAL" and Country_code=='US':
						pass
					else:
						row.append(quantity)
						row.append(Description)
						customers.append(row)

			
			
			output_str=""  #For Lables
					
			save_file = Main_Folder+"\\Outputs\\"+file.split('\\')[-1].split('.csv')[0]+" - Output.csv"
			with open(save_file, 'w', newline='', encoding='utf-8') as csvfile:
				spamwriter = csv.writer(csvfile, delimiter=',',
										quotechar='"')
				for row in customers:
					spamwriter.writerow(row)
					
			save_file = Main_Folder+"\\Outputs\\"+file.split('\\')[-1].split('.csv')[0]+" - StampsCom.csv"
			with open(save_file, 'w', newline='', encoding='utf-8') as csvfile:
				spamwriter = csv.writer(csvfile, delimiter=',',
										quotechar='"')
				headings_row = "Order ID (required),Order Date,Order Value,Requested Service,Ship To - Name,Ship To - Company,Ship To - Address 1,Ship To - Address 2,Ship To - Address 3,Ship To - State/Province,Ship To - City,Ship To - Postal Code,Ship To - Country,Ship To - Phone,Ship To - Email,Total Weight in Oz,Dimensions - Length,Dimensions - Width,Dimensions - Height,Notes - From Customer,Notes - Internal,Gift Wrap?,Gift Message"
				headings_row = headings_row.split(',')
				spamwriter.writerow(headings_row)
				
				for row in customers:
					Address1=row[4]
					if Address1:
						First=row[0]
						Last=row[2]
						Address1=row[4]
						Address2=row[5]
						Address3=row[6]
						City=row[7]
						State=row[8]
						Zip=row[9]
						Country_code=row[10]
						
						Description = file.split('\\')[-1].split(' - ')[0].split('USD')[1].strip()
						if Description=="One Fipsy FPGA" or Description=="One Pre-soldered Fipsy FPGA":
							quantity=1
						elif Description=="Two Fipsy Pack" or Description=="Pre-soldered Fipsy FPGA Two-Pack":
							quantity=2
						elif Description=="Fipsy FPGA Four-Pack":
							quantity=4
						elif Description=="Fipsy FPGA Ten-Pack" or Description=="Pre-soldered FPGA Ten-Pack":
							quantity=10
						else:
							raise Exception("Failed to match quantity to file name")
						
						
						if len(Country_code)==2:
							Country_name= pycountry.countries.get(alpha_2=Country_code).name
						else:
							Country_name=Country_code
						Phone=row[12]
						Email=row[14]
						ReferenceNumber=row[15]
						
						new_row=[]
						new_row.append(ReferenceNumber)
						new_row.append(ORDER_DATE)
						order_value = quantity*PER_ITEM_COST
						order_value = '${:,.2f}'.format(order_value)
						new_row.append(order_value)
						new_row.append(REQUESTED_SERVICE)
						new_row.append(First+" "+Last)
						Company = ""
						new_row.append(Company)
						new_row.append(Address1)
						new_row.append(Address2)
						new_row.append(Address3)
						new_row.append(State)
						new_row.append(City)
						new_row.append(Zip)
						new_row.append(Country_name)
						new_row.append(Phone)
						new_row.append(Email)
						new_row.append(ITEM_WEIGHT_OZ)
						new_row.append(ITEM_LENGTH)
						new_row.append(ITEM_WIDTH)
						new_row.append(ITEM_HEIGHT)
						new_row.append(shipping_details_dict[ReferenceNumber]) #Note from customer
						internal_note = ""
						new_row.append(internal_note)
						Gift_Wrap = "FALSE"
						new_row.append(Gift_Wrap)
						Gift_Message = ""
						new_row.append(Gift_Message)
						
						spamwriter.writerow(new_row)
					else:
						pass #An empty Address
					
					
					
					
			
	print("Now Done")