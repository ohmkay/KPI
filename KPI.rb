require 'csv'
require 'date'
require 'time'
require 'WriteExcel'
require_relative ('Users')


###############################################################################################
# This function takes CSV files and parses the data (line by line) into their appropriate class
###############################################################################################
def read_data(users_csv_path, cases_csv_path, tasks_csv_path, correspondence_csv_path, start_date, end_date)

	users = []

	#read string as path
	puts "Reading CSV data..." 
	users_csv = CSV.read(users_csv_path, {encoding: "UTF-8", headers: true, skip_blanks: true, header_converters: :symbol, converters: :all})
	cases_csv = CSV.read(cases_csv_path, {encoding: "UTF-8", headers: true, header_converters: :symbol, converters: :all})
	tasks_csv = CSV.read(tasks_csv_path, {encoding: "UTF-8", headers: true, header_converters: :symbol, converters: :all})

	correspondence_csv = []
	CSV.foreach(correspondence_csv_path, {encoding: "UTF-8", headers: true, header_converters: :symbol, converters: :all}) do |row|
		correspondence_date = Date.strptime(row[3], "%m/%d/%Y %H:%M") if !row[3].nil?
  		correspondence_csv << row if ((correspondence_date >= start_date) && (correspondence_date <= end_date))
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

		#simple check on days open before adding data to Case
		if !create_date.nil? && !close_date.nil? && status == -1 && close_date < end_date
			days_open = (close_date - create_date).to_i
		elsif create_date <= end_date
			days_open = (end_date - create_date).to_i
		end

		temp_case = Case.new(case_number, create_date, close_date, a_number, case_type, status, days_open)

		users.select do |user|
			user.add_cases(temp_case) if user.a_number == a_number
		end
	end

	#task list processing
	puts "Processing Tasks..."
	tasks_csv.each do |line|
		complete_date = line[:completedate]
		a_number = line[:taskowner]
		hours = line[:actualhours].to_f
		case_number = line[:caseno]

		complete_date = nil if complete_date.nil? || complete_date.empty?
		complete_date = Date.strptime(complete_date, "%m/%d/%Y %H:%M:%S") unless complete_date.nil?

		#puts "#{a_number} - #{case_number} - #{hours} - #{complete_date}" if (hours > 20) && (!complete_date.nil? && (complete_date >= start_date && complete_date <= end_date))
		
		#add task to user's task list if anumber matches the user
		users.select do |user|
			user.add_tasks(Struct::Task.new(complete_date, a_number, hours)) if user.a_number == a_number && (!complete_date.nil? && (complete_date >= start_date && complete_date <= end_date))
		end
	end

	#sort the correspondence by case number followed by entry date
	puts "Sorting Correspondence..."
	correspondence_csv.sort_by! { |x| [x[:caseno], x[:entrydate]]}

	#correspondence list processing
	puts "Processing Correspondence..."

	correspondence_csv.each do |line|
		case_number = line[:caseno]
		entry_date = line[:entrydate]
		comment = line[:comment].to_s
		entry_date = Date.strptime(entry_date, "%m/%d/%Y %H:%M") unless entry_date.nil?

		#only save correspondence to the case if it is in the user specified time period
		users.select do |user|
			user.cases.select do |ticket|
				ticket.add_correspondence(Struct::Correspondence.new(entry_date, comment)) if case_number == ticket.case_number
			end
		end

	end

	return users
end

####################################
# Writes column headers to excel doc
####################################
def write_headers_to_excel(workbook, worksheet, users, title, title2, average)
	column_format = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center', 
		:bg_color => 'gray', 
		:bold => 1)
	worksheet.write(0,0, "Team", column_format)
	worksheet.write(0,1, "Owner", column_format)
	if average == true
		worksheet.write(0,2, "Average_Open", column_format)
		worksheet.write(0,3, title, column_format)
		worksheet.write(0,4, title2, column_format)
	else
		worksheet.write(0,2, title, column_format)
		worksheet.write(0,3, title2, column_format)
	end
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
	when 'open_more_than_10'
		user.open_more_than_10
	end
end 


##########################
# Write General Statistics
##########################
 def write_worksheet(worksheet, users, user_case, calculate_avg, avg_type, tf1, tf2, cf1, cf2, mcf1, mcf2)
	
	row, sum = 2, 0
	team = users[0].team
	team_start = row
	alternate_count = false

	#writes user names and user teams with merged cells with formatting
	users.each do |user|

		if team != user.team
			if calculate_avg == false
				team_names = "A#{team_start}:A#{row-1}, #{team}"
				totals = "D#{team_start}:D#{row-1}, #{sum}"
			else
				team_names = "A#{team_start}:A#{row-1}, #{team}"
				totals = "E#{team_start}:E#{row-1}, #{sum}"
			end

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
			if calculate_avg == false
				worksheet.write_string(row-1, 1, user.name, cf1) 
				worksheet.write_number(row-1, 2, select_user_variable(user, user_case), cf1)
			else
				worksheet.write_string(row-1, 1, user.name, cf1) 

				if avg_type == 'closed'
					worksheet.write_number(row-1, 2, user.avg_closed[:average], cf1)
				elsif avg_type == 'open'
					worksheet.write_number(row-1, 2, user.avg_open[:average], cf1)
				end

				worksheet.write_number(row-1, 3, select_user_variable(user, user_case), cf1)
			end
		else
			if calculate_avg == false
				worksheet.write_string(row-1, 1, user.name, cf2)
				worksheet.write_number(row-1, 2, select_user_variable(user, user_case), cf2)
			else
				worksheet.write_string(row-1, 1, user.name, cf2)

				if avg_type == 'closed'
					worksheet.write_number(row-1, 2, user.avg_closed[:average], cf2)
				elsif avg_type == 'open'
					worksheet.write_number(row-1, 2, user.avg_open[:average], cf2)
				end

				worksheet.write_number(row-1, 3, select_user_variable(user, user_case), cf2)
			end
		end

		team = user.team
		row += 1
		sum += select_user_variable(user, user_case) if !select_user_variable(user, user_case).nil?
	end

	#processing for last team
	if calculate_avg == false
		team_names = "A#{team_start}:A#{row-1}, #{team}"
		totals = "D#{team_start}:D#{row-1}, #{sum}"
	else
		team_names = "A#{team_start}:A#{row-1}, #{team}"
		totals = "E#{team_start}:E#{row-1}, #{sum}"
	end
			
	if alternate_count == true
		worksheet.merge_range(team_names, team, tf1)
		worksheet.merge_range(totals, sum, mcf1)
	else
		worksheet.merge_range(team_names, team, tf2)
		worksheet.merge_range(totals, sum, mcf2)
	end
	
end

def calculate_totals(worksheet, workbook, users)
	column_format_totals = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center', 
		:bg_color => 'gray', 
		:bold => 1)
	worksheet.write(0,0, "Created Cases", column_format_totals)
	worksheet.write(0,1, "Closed Cases", column_format_totals)
	worksheet.write(0,2, "Open Cases", column_format_totals)
	worksheet.write(0,3, "Hours", column_format_totals)
	worksheet.write(0,4, "Average Closed Cases", column_format_totals)
	worksheet.write(0,5, "Average Open Cases", column_format_totals)

	created_total = closed_total = open_total = hours_total = 0
	avg_closed_count = avg_open_count = avg_closed_days_total = avg_open_days_total = 0

	users.each do |user|
		 created_total += user.created_cases
		 closed_total += user.closed_cases
		 open_total += user.open_cases
		 hours_total += user.hours_total

		 avg_closed_count += user.avg_closed[:count]
		 avg_closed_days_total += user.avg_closed[:days_total]

		 avg_open_count += user.avg_open[:count]
		 avg_open_days_total += user.avg_open[:days_total]
	end	

	worksheet.write(1,0,created_total)
	worksheet.write(1,1,closed_total)
	worksheet.write(1,2,open_total)
	worksheet.write(1,3,hours_total)
	worksheet.write(1,4,(avg_closed_days_total/avg_closed_count))
	worksheet.write(1,5,(avg_open_days_total/avg_open_count))
end

###############
# Main Function
###############
def run_forrest_run

	#Change these dates
	start_date = Date.new(2016,04,01)
	end_date = Date.new(2016,04,30)

	#set csv paths
	users_csv = 'C:\Users\A5NB3ZZ\Documents\Projects\2016\5 - May\KPI\userList.txt'
	cases_csv = 'C:\Users\A5NB3ZZ\Documents\Projects\2016\5 - May\KPI\Cases.txt'
	tasks_csv = 'C:\Users\A5NB3ZZ\Documents\Projects\2016\5 - May\KPI\Tasks.txt'
	correspondence_csv = 'C:\Users\A5NB3ZZ\Documents\Projects\2016\5 - May\KPI\Correspondence2.txt'

	#number of days apart to check for case inactivity
	days_inactive = 7

	#read in data from CSV into array users of User class
	users = read_data(users_csv, cases_csv, tasks_csv, correspondence_csv, start_date, end_date)
	
	#sort users based on team
	users.sort! {|x, y| x.team <=> y.team}

	#generate stats for each user via the user class
	users.each {|x| x.generate_user_statistics(start_date, end_date, days_inactive)}

	#change this for the output filename/path
	workbook = WriteExcel.new('C:\Users\A5NB3ZZ\Documents\Projects\2016\5 - May\KPI\testApril.xls')
	tf1, tf2, cf1, cf2, mcf1, mcf2 = generate_styles(workbook)


	###################
	#Worksheet Creation
	###################
	monthly_totals = workbook.add_worksheet('Monthly Summary')
	monthly_totals.set_column('A:F', 20)
	calculate_totals(monthly_totals, workbook, users)

	hours_worksheet = workbook.add_worksheet('Hours')
	hours_worksheet.set_column('A:D', 20)
	write_headers_to_excel(workbook, hours_worksheet, users, "Hours Sum", "Hours Sum Per Team", false)
	write_worksheet(hours_worksheet, users, 'hours', false, 'none', tf1, tf2, cf1, cf2, mcf1, mcf2)

	open_worksheet = workbook.add_worksheet('Open')
	open_worksheet.set_column('A:E', 20)
	write_headers_to_excel(workbook, open_worksheet, users, "Open Sum", "Open Sum Per Team", true)
	write_worksheet(open_worksheet, users, 'open', true, 'open', tf1, tf2, cf1, cf2, mcf1, mcf2)

	created_worksheet = workbook.add_worksheet('Created')
	created_worksheet.set_column('A:D', 20)
	write_headers_to_excel(workbook, created_worksheet, users, "Created Sum", "Created Sum Per Team", false)
	write_worksheet(created_worksheet, users, 'created', false, 'none', tf1, tf2, cf1, cf2, mcf1, mcf2)

	closed_worksheet = workbook.add_worksheet('Closed')
	closed_worksheet.set_column('A:E', 20)
	write_headers_to_excel(workbook, closed_worksheet, users, "Closed Sum", "Closed Sum Per Team", true)
	write_worksheet(closed_worksheet, users, 'closed', true, 'closed', tf1, tf2, cf1, cf2, mcf1, mcf2)

	inactive_worksheet = workbook.add_worksheet('Inactive')
	inactive_worksheet.set_column('A:D', 20)
	write_headers_to_excel(workbook, inactive_worksheet, users, "Inactive Sum", "Inactive Sum Per Team", false)
	write_worksheet(inactive_worksheet, users, 'inactive', false, 'none', tf1, tf2, cf1, cf2, mcf1, mcf2)

	open_more_than_10 = workbook.add_worksheet('OpenMoreThan10')
	open_more_than_10.set_column('A:D', 20)
	write_headers_to_excel(workbook, open_more_than_10, users, "Open > 10 Sum", "Sum Per Team", false)
	write_worksheet(open_more_than_10, users, 'open_more_than_10', false, 'none', tf1, tf2, cf1, cf2, mcf1, mcf2)

	workbook.close
end

run_forrest_run()