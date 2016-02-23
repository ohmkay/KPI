class User
	attr_accessor :user_id, :a_number, :name, :team, :cases, :tasks

	def initialize(user_id, a_number, name, team)
		@user_id = user_id
		@a_number = a_number
		@name = name
		@team = team
		@cases = []
		@tasks = []
	end

	def add_cases(ticket)
		cases.push(ticket)
	end

	def add_tasks(task)
		tasks.push(task)
	end

	def count_open_cases(start_date)
		open_cases = 0

		cases.each do |ticket|
			open_cases += 1 if (ticket.create_date != nil && ticket.create_date >= start_date)
		end
		open_cases
	end

	def count_closed_cases(start_date, end_date)
		closed_cases = 0

		cases.each do |ticket|
			closed_cases += 1 if (ticket.close_date != nil) && (ticket.close_date <= end_date && ticket.close_date >= start_date)	
		end
		closed_cases
	end

	def count_hours(start_date, end_date)
		hours_total = 0

		tasks.each do |task_row|
			hours_total += task_row.hours if (task_row.hours != nil) && (task_row.complete_date <= end_date) && (task_row.complete_date >= start_date)
		end
		hours_total
	end

end

class Case
	attr_accessor :create_date, :close_date, :a_number

	def initialize(create_date, close_date, a_number)
		@create_date = create_date
		@close_date = close_date
		@a_number = a_number
	end
end

class Task
	attr_accessor :complete_date, :a_number, :hours

	def initialize(complete_date, a_number, hours)
		@complete_date = complete_date
		@a_number = a_number
		@hours = hours
	end
end