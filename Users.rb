require_relative ('Case')

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

	def generate_user_statistics(start_date, end_date, days_inactive)
		cases.each do |ticket|
			#generate ticket level statistics via the Case class
			ticket.generate_case_statistics(start_date, end_date, days_inactive)

				if ticket.closed_status == true
					@closed_cases += 1
					@avg_closed[:days_total] += ticket.days_open if !ticket.days_open.nil?
					@avg_closed[:count] += 1
				end

				if ticket.open_status == true
					@open_cases += 1
					@avg_open[:days_total] += ticket.days_open if !ticket.days_open.nil?
					@avg_open[:count] += 1
				end

				@created_cases += 1 if ticket.created_status == true
				@inactive_cases += 1 if ticket.inactive_status == true
				puts "#{@a_number} - #{ticket.case_number} - #{ticket.create_date} - #{ticket.close_date}" if ticket.inactive_status == true && @a_number == "USAC\\A6689ZZ"
				@open_more_than_10 += 1 if ticket.openmorethan10_status == true
		end

		#calculating averages and average totals for open and closed tickets
		@avg_closed[:average] = @avg_closed[:days_total] / @avg_closed[:count] if @avg_closed[:count] != 0
		@avg_open[:average] = @avg_open[:days_total] / @avg_open[:count] if @avg_open[:count] != 0

		#calculate hours based on tasks
		tasks.each do |task|
			@hours_total += task[:hours] if (!task[:hours].nil? && !task[:complete_date].nil?) && 
			(task[:complete_date] <= end_date) && (task[:complete_date] >= start_date)
		end
	end

end

Struct.new("Task", :complete_date, :a_number, :hours)