require 'csv'
require 'date'
require 'time'
require 'WriteExcel'
require_relative ('Users')

#####################################################################################
# This function takes 3 CSV files and parses the data (line by line) into their 
# appropriate class.  It utilizes the global variable $users to hold all data
#####################################################################################
def read_data(users_csv, cases_csv, tasks_csv)

	users = []

	#read string as path
	users_csv = CSV.read(users_csv)
	cases_csv = CSV.read(cases_csv)
	tasks_csv = CSV.read(tasks_csv)

	#remove headers from CSV
	users_csv.slice!(0)
	cases_csv.slice!(0)
	tasks_csv.slice!(0)

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
		create_date = line[3]
		create_date = Date.strptime(create_date, "%m/%d/%Y %H:%M:%S") if !create_date.nil?

		close_date = line[5]
		close_date = Date.strptime(close_date, "%m/%d/%Y %H:%M:%S") if !close_date.nil?
		
		a_number = line[1]

		temp_case = Case.new(create_date, close_date, a_number)
		
		users.select do |user|
			user.add_cases(temp_case) if user.a_number == a_number
		end
	end

	#task list processing
	tasks_csv.each do |line|
		complete_date = line[6]
		complete_date = Date.strptime(complete_date, "%m/%d/%Y %H:%M:%S") if !complete_date.nil?

	 	a_number = line[3]
		hours = line[11].to_f

		temp_task = Task.new(complete_date, a_number, hours)
		
		users.select do |user|
			user.add_tasks(temp_task) if user.a_number == a_number
		end
	end

	return users

end

###############################
# Generate Statistics
###############################
def write_to_excel(users, start_date, end_date)
	workbook = WriteExcel.new('../text.xlsx')

	team_summary_worksheet = workbook.add_worksheet

	write_data(worksheet,1,10, 19019)
	
	workbook.close
end

###############################
# Main Function
###############################
def run_forrest_run

	#Change these dates
	start_date = Date.new(2016,01,01)
	end_date = Date.new(2016,01,31)

	#read in data
	users_csv = 'C:\Users\Caleb\Documents\Ruby\CSVs\userList.csv'
	cases_csv = 'C:\Users\Caleb\Documents\Ruby\CSVs\CogentCase.csv'
	tasks_csv = 'C:\Users\Caleb\Documents\Ruby\CSVs\CaseTask_Hours2.csv'

	#read in data from CSV into array users of User class
	users = read_data(users_csv, cases_csv, tasks_csv)

	users.each {|user| puts user.team}

	users.sort! { |x, y| x.team <=> y.team}
	puts '-' * 25

	users.each do |user| 
		puts user.team 
		puts user.name
	end


	users[5].generate_statistics(start_date, end_date)
	puts users[5].created_cases

	###############################
	# Testing and debugging section
	###############################
	#users.each do |user|
	#	total_tickets = 0
	#	puts user.name
	#	user.cases.each do |ticket|
	#		total_tickets += 1
	#	end
	#	puts "Total tickets: #{total_tickets}"
	#	puts '-' * 25
	#end



end

run_forrest_run()