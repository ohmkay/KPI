require 'csv'
require 'date'
require 'time'
require 'WriteExcel'
require_relative ('Users')

#####################################################################################
# This function takes CSV files and parses the data (line by line) into their 
# appropriate class.  It utilizes the global variable $users to hold all data
#####################################################################################
def read_data(users_csv, cases_csv, tasks_csv, correspondence_csv, start_date, end_date)

	users = []

	#read string as path
	users_csv = CSV.read(users_csv)
	cases_csv = CSV.read(cases_csv)
	tasks_csv = CSV.read(tasks_csv)
	correspondence_csv = CSV.read(correspondence_csv)

	#remove headers from CSV
	users_csv.slice!(0)
	cases_csv.slice!(0)
	tasks_csv.slice!(0)
	correspondence_csv.slice!(0)

	#user list processing
	users_csv.each do |line|
	 	user_id = line[0]
	 	team = line[2]
		name = line[1]
		a_number = line[3]

		temp_user = User.new(user_id, a_number, name, team)
		users.push(temp_user)
	end

	#case list processing
	cases_csv.each do |line|
		status = line[2]
		create_date = Date.strptime(line[3], "%m/%d/%Y %H:%M:%S") if !line[3].nil?
		close_date = Date.strptime(line[5], "%m/%d/%Y %H:%M:%S") if !line[5].nil?
		a_number = line[1]

		users.select do |user|
			user.add_cases(Struct::Case.new(create_date, close_date, a_number)) if (user.a_number == a_number) && (status != -1 && (close_date.nil? || (close_date > start_date)))
		end
	end

	#task list processing
	tasks_csv.each do |line|
		complete_date = Date.strptime(line[6], "%m/%d/%Y %H:%M:%S") if !line[6].nil?
	 	a_number = line[3]
		hours = line[11].to_f
		
		users.select do |user|
			user.add_tasks(Struct::Task.new(complete_date, a_number, hours)) if (user.a_number == a_number) && (complete_date.nil? || (complete_date >= start_date && complete_date <= end_date))
		end
	end
	return users
end

####################################
# Writes column headers to excel doc
####################################
def write_headers_to_excel(workbook, worksheet, users, title, title2)
	column_format = workbook.add_format(:valign  => 'vcenter', :align   => 'center', :bg_color => 'green')
	worksheet.write(0,0, "Team", column_format)
	worksheet.write(0,1, "Owner", column_format)
	worksheet.write(0,2, title, column_format)
	worksheet.write(0,3, title2, column_format)
end

#######################################################
# Generate worksheet styles for alternating team colors
#######################################################
def generate_styles(workbook)
	title_format_1 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:rotation => '90', 
		:bold => true,
		:bg_color => 'blue'	
	)
	title_format_2 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:rotation => '90', 
		:bold => true,
		:bg_color => 'red'	
	)
	cell_format_1 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:bg_color => 'blue'
	)
	cell_format_2 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:bg_color => 'red'
	)
	merged_cell_format_1 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:bg_color => 'blue'
	)
	merged_cell_format_2 = workbook.add_format(
		:valign  => 'vcenter', 
		:align   => 'center',
		:bg_color => 'red'
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
	end
end 
#####################
# Generate Statistics
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
				worksheet.merge_range(totals, sum, cf1)
			else
				worksheet.merge_range(team_names, team, tf2)
				worksheet.merge_range(totals, sum, cf2)
			end

			team_start = row
			sum = 0
			alternate_count = !alternate_count
		end

		if alternate_count == true
			worksheet.write(row-1, 1, user.name, cf1) 
			worksheet.write(row-1, 2, select_user_variable(user, user_case), mcf1)
		else
			worksheet.write(row-1, 1, user.name, cf2)
			worksheet.write(row-1, 2, select_user_variable(user, user_case), mcf2)
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
		worksheet.merge_range(totals, sum, cf1)
	else
		worksheet.merge_range(team_names, team, tf2)
		worksheet.merge_range(totals, sum, cf2)
	end
	
end


###############
# Main Function
###############
def run_forrest_run

	#Change these dates
	start_date = Date.new(2016,01,01)
	end_date = Date.new(2016,01,31)

	#set csv paths
	users_csv = 'C:\Users\caleb\Dropbox\KPICSV\userList.csv'
	cases_csv = 'C:\Users\caleb\Dropbox\KPICSV\CogentCase.csv'
	tasks_csv = 'C:\Users\caleb\Dropbox\KPICSV\CaseTask_Hours2.csv'
	correspondence_csv = 'C:\Users\caleb\Dropbox\KPICSV\CaseCorrespondence.csv'

	#read in data from CSV into array users of User class
	users = read_data(users_csv, cases_csv, tasks_csv, correspondence_csv, start_date, end_date)
	
	users.sort! {|x, y| x.team <=> y.team}
	users.each {|x| x.generate_statistics(start_date, end_date)}

	#write_to_excel(users)

	workbook = WriteExcel.new('./test.xls')
	tf1, tf2, cf1, cf2, mcf1, mcf2 = generate_styles(workbook)

	hours_worksheet = workbook.add_worksheet('Hours')
	write_headers_to_excel(workbook, hours_worksheet, users, "Hours Sum", "Hours Sum Per Team")
	write_worksheet(hours_worksheet, users, 'hours', tf1, tf2, cf1, cf2, mcf1, mcf2)

	open_worksheet = workbook.add_worksheet('Open_Cases')
	write_headers_to_excel(workbook, open_worksheet, users, "Open Sum", "Open Sum Per Team")
	write_worksheet(open_worksheet, users, 'open', tf1, tf2, cf1, cf2, mcf1, mcf2)

	closed_worksheet = workbook.add_worksheet('Closed_Cases')
	write_headers_to_excel(workbook, closed_worksheet, users, "Closed Sum", "Closed Sum Per Team")
	write_worksheet(closed_worksheet, users, 'closed', tf1, tf2, cf1, cf2, mcf1, mcf2)

	inactive_worksheet = workbook.add_worksheet('Inactive')
	write_headers_to_excel(workbook, closed_worksheet, users, "Inactive Sum", "Inactive Sum Per Team")
	

	workbook.close
end

run_forrest_run()