class User
	attr_accessor :user_id, :a_number, :name, :team, :cases, :tasks, 
	:created_cases, :closed_cases, :open_cases, :hours_total, :inactive_cases, :open_more_than_10, :avg_open, :avg_closed

	def initialize(user_id, a_number, name, team)
		@user_id = user_id
		@a_number = a_number
		@name = name
		@team = team
		@cases = []
		@tasks = []
		@avg_open = {:days_total => 0, :count => 0, :average => 0}
		@avg_closed = {:days_total => 0, :count => 0, :average => 0}
		@created_cases = @closed_cases = @open_cases = 
		@hours_total = @inactive_cases = @open_more_than_10 = 0
	end

	def add_cases(ticket)
		cases.push(ticket)
	end

	def add_tasks(task)
		tasks.push(task)
	end

	def generate_statistics(start_date, end_date)
		cases.each do |ticket|
			if (!ticket[:close_date].nil? && (ticket[:close_date] <= end_date && ticket[:close_date] >= start_date))
				@closed_cases += 1 
				@avg_closed[:days_total] += ticket.days_open
				@avg_closed[:count] += 1
			end

			if ticket[:close_date].nil? || (ticket[:close_date] >= end_date && ticket[:create_date] <= end_date)
				@open_cases += 1 
				@avg_open[:days_total] += ticket.days_open if !ticket.days_open.nil?
				@avg_open[:count] += 1
			end
			
			@open_more_than_10 += 1 if ticket[:close_date].nil? && (ticket[:create_date] + 10) < end_date
			@inactive_cases += 1 if (ticket[:inactive] == true)
		end
		@avg_closed[:average] = @avg_closed[:days_total] / @avg_closed[:count] if @avg_closed[:count] != 0
		@avg_open[:average] = @avg_open[:days_total] / @avg_open[:count] if @avg_open[:count] != 0

		tasks.each do |task|
			@hours_total += task[:hours] if (!task[:hours].nil? && !task[:complete_date].nil?) && 
			(task[:complete_date] <= end_date) && (task[:complete_date] >= start_date)
		end
	end

end


Struct.new("Case", :case_number, :create_date, :close_date, :a_number, :case_type, :status, :creator, :days_open, :inactive)
Struct.new("Task", :complete_date, :a_number, :hours)