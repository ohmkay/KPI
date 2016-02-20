require 'date'
require 'time'
require_relative ('Users')

#####################################################################################
# This function takes 3 CSV files and parses the data (line by line) into their 
# appropriate class.  It utilizes the global variable $users to hold all data
#####################################################################################
def read_data(users_csv, cases_csv, tasks_csv, start_date, end_date)
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
		$users.push(temp_user)
	end

	#case list processing
	cases_csv.each do |line|
		create_date = line[3]
		create_date = Date.strptime(create_date, "%m/%d/%Y %H:%M:%S") if !create_date.nil?

		close_date = line[5]
		close_date = Date.strptime(close_date, "%m/%d/%Y %H:%M:%S") if !close_date.nil?
		
		a_number = line[1]

		temp_case = Case.new(create_date, close_date, a_number)
		
		$users.select do |user|
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
		
		$users.select do |user|
			user.add_tasks(temp_task) if user.a_number == a_number
		end
	end

end