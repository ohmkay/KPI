class User
	attr_accessor :user_id, :a_number, :name, :team, :cases, :tasks, 
	:created_cases, :closed_cases, :open_cases, :hours_total, :inactive_cases

	def initialize(user_id, a_number, name, team)
		@user_id = user_id
		@a_number = a_number
		@name = name
		@team = team
		@cases = []
		@tasks = []
		@created_cases = @closed_cases = @open_cases = 
		@hours_total = @inactive_cases = 0
	end

	def add_cases(ticket)
		cases.push(ticket)
	end

	def add_tasks(task)
		tasks.push(task)
	end

	def generate_statistics(start_date, end_date)
		cases.each do |ticket|
			@created_cases += 1 if (ticket[:create_date].nil? && ticket[:create_date] >= start_date)
			@closed_cases += 1 if (!ticket[:close_date].nil? && (ticket[:close_date] <= end_date && ticket[:close_date] >= start_date))	
			@open_cases += 1 if (ticket[:close_date].nil? || ticket[:close_date] <= end_date)
			@inactive_cases += 1 if (ticket[:inactive] == true)
			#puts "#{ticket[:case_number]}" if (ticket[:inactive] == true)
		end


		tasks.each do |task|
			@hours_total += task[:hours] if (!task[:hours].nil? && !task[:complete_date].nil?) && 
			(task[:complete_date] <= end_date) && (task[:complete_date] >= start_date)
		end
	end

end

Struct.new("Case", :case_number, :create_date, :close_date, :a_number, :inactive)
Struct.new("Task", :complete_date, :a_number, :hours)