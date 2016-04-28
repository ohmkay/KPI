require 'csv'
require 'date'
require 'time'
require 'WriteExcel'
require_relative ('Users')


#####################################################################################
# This function takes CSV files and parses the data (line by line) into their 
# appropriate class.  It utilizes the global variable $users to hold all data
#####################################################################################
def read_data(users_csv_path, cases_csv_path, tasks_csv_path, correspondence_csv_path, start_date, end_date)

	users = []

	#read string as path
	puts "Reading CSV data..." 
	users_csv = CSV.read(users_csv_path, {encoding: "UTF-8", headers: true, header_converters: :symbol, converters: :all})
	cases_csv = CSV.read(cases_csv_path, {encoding: "UTF-8", headers: true, header_converters: :symbol, converters: :all})
	tasks_csv = CSV.read(tasks_csv_path, {encoding: "UTF-8", headers: true, header_converters: :symbol, converters: :all})

	correspondence_csv = []
	CSV.foreach(correspondence_csv_path, {encoding: "UTF-8", headers: true, header_converters: :symbol, converters: :all}) do |row|
  		correspondence_csv << row if Date.strptime(row[3], "%m/%d/%Y %H:%M") >= start_date 
	end

	#user list processing
	puts "Processing Users..."
	users_csv.each do |line|
	 	user_id = line[:userid]
	 	team = line[:team]
		name = line[:name]
		a_number = line[:a_number]

		temp_user = User.new(user_id, a_number, name, team)
		users.push(temp_user)
	end

	#case list processing
	puts "Processing Cases..."
	cases_csv.each do |line|
		case_number = line[:caseno]
		status = line[:status]
		create_date = line[:createdate]
		close_date = line[:closedate]
		a_number = line[:owner]
		case_type = line[:casetype]

		#only read in Support cases
		next if case_type != 'Support'

		create_date = nil if create_date.nil? || create_date.empty?
		create_date = Date.strptime(create_date, "%m/%d/%Y %H:%M:%S") unless create_date.nil?
		close_date = nil if close_date.nil? || close_date.empty?
		close_date = Date.strptime(close_date, "%m/%d/%Y %H:%M:%S") unless close_date.nil?

		#add case to user's case list if anumber matches the user
		users.select do |user|
			user.add_cases(Struct::Case.new(case_number, create_date, close_date, a_number, case_type)) if user.a_number == a_number
		end
	end

	#task list processing
	puts "Processing Tasks..."
	tasks_csv.each do |line|
		complete_date = line[:completedate]
		a_number = line[:taskowner]
		hours = line[:actualhours].to_f

		complete_date = nil if complete_date.nil? || complete_date.empty?
		complete_date = Date.strptime(complete_date, "%m/%d/%Y %H:%M:%S") unless complete_date.nil?
		
		#add task to user's task list if anumber matches the user
		users.select do |user|
			user.add_tasks(Struct::Task.new(complete_date, a_number, hours)) if user.a_number == a_number && complete_date.nil? || (complete_date >= start_date && complete_date <= end_date)
		end
	end

	#sort the correspondence by case number followed by entry date
	puts "Sorting Correspondence..."
	correspondence_csv.sort_by! { |x| [x[:caseno], x[:entrydate]]}

	previous_case_number = nil
	previous_entry_date = start_date

	#correspondence list processing
	puts "Processing Correspondence..."
	correspondence_csv.each do |line|
		case_number = line[:caseno]
		entry_date = line[:entrydate]
		entry_date = Date.strptime(entry_date, "%m/%d/%Y %H:%M") unless entry_date.nil?

		#changes date check to start date if case changes
		if previous_case_number != case_number

			#checks if last correspondence has gap between itself and last day of the period
			if end_date > previous_entry_date + 7
				users.select do |user|
					user.cases.select do |ticket|
						if previous_case_number == ticket.case_number
							ticket.inactive = true
							#puts "YES #2 #{ticket.case_number} - #{end_date} - #{previous_entry_date}"
						end
					end
				end
			end

			previous_entry_date = start_date
		end

		#confirms if inactivity is in given date range
		if ((entry_date >= start_date) && (entry_date <= end_date))
			#checks if case number is the same as the previous case number
			if ((case_number == previous_case_number) && !previous_case_number.nil?)
				#checks if the dates are greater than 7 days apart
				if ((entry_date > previous_entry_date + 7) && !previous_entry_date.nil?)
					#finds ticket to mark inactive 
					users.select do |user|
						user.cases.select do |ticket|
							if (case_number == ticket.case_number)
								ticket.inactive = true 
								#puts "YES #{ticket.case_number} - #{entry_date} - #{previous_entry_date}"
							end
						end
					end
				end
			end
		end

		previous_case_number = case_number
		previous_entry_date = entry_date
	end

	return users
end

####################################
# Writes column headers to excel doc
####################################
def write_headers_to_excel(workbook, worksheet, users, title, title2)
	column_format = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center', 
		:bg_color => 'gray', 
		:bold => 1)
	worksheet.write(0,0, "Team", column_format)
	worksheet.write(0,1, "Owner", column_format)
	worksheet.write(0,2, title, column_format)
	worksheet.write(0,3, title2, column_format)
end

#######################################################
# Generate worksheet styles for alternating team colors
#######################################################
def generate_styles(workbook)
	color_blue = workbook.set_custom_color(40, '#33CCFF')
	color_orange = workbook.set_custom_color(41, '#FFCC33')

	title_format_1 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:rotation => '90', 
		:bold => 1,
		:bg_color => color_blue,
		:set_color => 'black'
	)
	title_format_2 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:rotation => '90', 
		:bold => 1,
		:bg_color => color_orange,
		:set_color => 'black'
	)
	cell_format_1 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:bg_color => color_blue,
		:set_color => 'black'
	)
	cell_format_2 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:bg_color => color_orange,
		:set_color => 'black'
	)
	merged_cell_format_1 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:bg_color => color_blue,
		:set_color => 'black'
	)
	merged_cell_format_2 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:bg_color => color_orange,
		:set_color => 'black'
	)

	return title_format_1, title_format_2, cell_format_1, cell_format_2, 
		   merged_cell_format_1, merged_cell_format_2
end


####################################################
# changes which variable is selected from user class
####################################################
def select_user_variable(user, user_case)
	case user_case
	when 'hours'
		user.hours_total
	when 'created'
		user.created_cases
	when 'open'
		user.open_cases
	when 'closed'
		user.closed_cases
	when 'inactive'
		user.inactive_cases
	end
end 


#####################
# Generate Statistics
#####################
 def write_worksheet(worksheet, users, user_case, tf1, tf2, cf1, cf2, mcf1, mcf2)
	
	row, sum = 2, 0
	team = users[0].team
	team_start = row
	alternate_count = false

	#writes user names and user teams with merged cells with formatting
	users.each do |user|

		if team != user.team
			team_names = "A#{team_start}:A#{row-1}, #{team}"
			totals = "D#{team_start}:D#{row-1}, #{sum}"

			if alternate_count == true
				worksheet.merge_range(team_names, team, tf1)
				worksheet.merge_range(totals, sum, mcf1)
			else
				worksheet.merge_range(team_names, team, tf2)
				worksheet.merge_range(totals, sum, mcf2)
			end

			team_start = row
			sum = 0
			alternate_count = !alternate_count
		end

		if alternate_count == true
			worksheet.write_string(row-1, 1, user.name, cf1) 
			worksheet.write_number(row-1, 2, select_user_variable(user, user_case), cf1)
		else
			worksheet.write_string(row-1, 1, user.name, cf2)
			worksheet.write_number(row-1, 2, select_user_variable(user, user_case), cf2)
		end

		team = user.team
		row += 1
		sum += select_user_variable(user, user_case) if !select_user_variable(user, user_case).nil?
	end

	#processing for last team
	team_names = "A#{team_start}:A#{row-1}, #{team}"
	totals = "D#{team_start}:D#{row-1}, #{sum}"
			
	if alternate_count == true
		worksheet.merge_range(team_names, team, tf1)
		worksheet.merge_range(totals, sum, mcf1)
	else
		worksheet.merge_range(team_names, team, tf2)
		worksheet.merge_range(totals, sum, mcf2)
	end
	
end


###############
# Main Function
###############
def run_forrest_run

	#Change these dates
	start_date = Date.new(2016,03,01)
	end_date = Date.new(2016,03,31)

	#set csv paths
	users_csv = 'C:\Users\caleb\Documents\Ruby\newKPIMarch\userList.csv'
	cases_csv = 'C:\Users\caleb\Documents\Ruby\newKPIMarch\Cases.csv'
	tasks_csv = 'C:\Users\caleb\Documents\Ruby\newKPIMarch\Tasks.csv'
	correspondence_csv = 'C:\Users\caleb\Documents\Ruby\newKPIMarch\Correspondence.csv'

	#read in data from CSV into array users of User class
	users = read_data(users_csv, cases_csv, tasks_csv, correspondence_csv, start_date, end_date)

	#TESTING
	#users.select do |user|
	#	if user.a_number == 'USAC\\A5NB3ZZ'
	#		user.cases.each do |ticket|
	#			#puts "#{ticket[:close_date]} ----  #{end_date}"
	#			puts "#{ticket.a_number} - #{ticket.case_number} - #{ticket.create_date} - #{ticket.close_date} - #{ticket.case_type}" if (!ticket[:close_date].nil? && (ticket[:close_date] <= end_date && ticket[:close_date] >= start_date))	
	#		end
	#	end
	#end
	
	#sort users based on team
	users.sort! {|x, y| x.team <=> y.team}

	#generate stats for each user via the user class
	users.each {|x| x.generate_statistics(start_date, end_date)}

	#change this for the output filename/path
	workbook = WriteExcel.new('C:\Users\caleb\Documents\Ruby\newKPIMarch\output.xls')
	tf1, tf2, cf1, cf2, mcf1, mcf2 = generate_styles(workbook)


	###################
	#Worksheet Creation
	###################
	hours_worksheet = workbook.add_worksheet('Hours')
	hours_worksheet.set_column('B:B', 20)
	write_headers_to_excel(workbook, hours_worksheet, users, "Hours Sum", "Hours Sum Per Team")
	write_worksheet(hours_worksheet, users, 'hours', tf1, tf2, cf1, cf2, mcf1, mcf2)

	open_worksheet = workbook.add_worksheet('Open_Cases')
	open_worksheet.set_column('B:B', 20)
	write_headers_to_excel(workbook, open_worksheet, users, "Open Sum", "Open Sum Per Team")
	write_worksheet(open_worksheet, users, 'open', tf1, tf2, cf1, cf2, mcf1, mcf2)

	created_worksheet = workbook.add_worksheet('Created_Cases')
	created_worksheet.set_column('B:B', 20)
	write_headers_to_excel(workbook, created_worksheet, users, "Created Sum", "Created Sum Per Team")
	write_worksheet(created_worksheet, users, 'created', tf1, tf2, cf1, cf2, mcf1, mcf2)

	closed_worksheet = workbook.add_worksheet('Closed_Cases')
	closed_worksheet.set_column('B:B', 20)
	write_headers_to_excel(workbook, closed_worksheet, users, "Closed Sum", "Closed Sum Per Team")
	write_worksheet(closed_worksheet, users, 'closed', tf1, tf2, cf1, cf2, mcf1, mcf2)

	inactive_worksheet = workbook.add_worksheet('Inactive')
	inactive_worksheet.set_column('B:B', 20)
	write_headers_to_excel(workbook, inactive_worksheet, users, "Inactive Sum", "Inactive Sum Per Team")
	write_worksheet(inactive_worksheet, users, 'inactive', tf1, tf2, cf1, cf2, mcf1, mcf2)
	

	workbook.close
end

run_forrest_run()