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