class Case
	attr_accessor :case_number, :create_date, :close_date, :a_number, :case_type, :status, :creator, :days_open, :correspondence,
				  :open_status, :closed_status, :created_status, :inactive_status, :openmorethan10_status

	def initialize(case_number, create_date, close_date, a_number, case_type, status, days_open)
		@case_number = case_number
		@create_date = create_date
		@close_date = close_date
		@a_number = a_number
		@case_type = case_type
		@status = status
		@days_open = days_open
		@inactive_status = false
		@open_status = false
		@closed_status = false
		@created_status = false
		@openmorethan10_status = false
		@correspondence = []
	end

	def add_correspondence(comment)
		@correspondence.push(comment)
	end

	def check_inactive(start_date, end_date, days_inactive)
			previous_entry_date = nil
			@correspondence.each_with_index do |item, index|
				
				#determines if case was reassigned and marks inactive if true
				if (!item[:comment].nil?) && (item[:comment].include? "Case reassigned from")
					@inactive_status = false
					break
				end

				#completes check against start date for first item in correspondence
				if index == 0
					@inactive_status = true if ((item[:entry_date] > (start_date + days_inactive)) && (item[:entry_date] > (@create_date + days_inactive)))
				end

				#completes check against end date for last item in correspondence
				if index == @correspondence.size - 1
					@inactive_status = true if ((end_date > (item[:entry_date] + days_inactive)) && (@close_date.nil? || (@close_date > (item[:entry_date] + days_inactive))))
				end

				#compares previous iteration against current to determine inactivity
				@inactive_status = true if (!previous_entry_date.nil? && (item[:entry_date] > (previous_entry_date + days_inactive)))

				#sets current entry to previous to check for in next iteration
				previous_entry_date = item[:entry_date]

				#stops checking for inactivity if already marked true; breaks loop
				break if @inactive_status == true
			end

			#sets inactivity to true if there are no correspondence for the period being checked
			@inactive_status = true if @correspondence.length == 0
		#end
	end

	def check_open(start_date, end_date)
		@open_status = true if ((@create_date <= end_date) && (@close_date.nil? || (@close_date > end_date)))
	end

	def check_closed(start_date, end_date)
		@closed_status = true if (!@close_date.nil? && (@close_date <= end_date) && (@close_date >= start_date))		
	end

	def check_created(start_date, end_date)
		@created_status = true if (!@create_date.nil? && (@create_date >= start_date) && (@create_date <= end_date))
	end

	def check_openmorethan10(start_date, end_date)
		@openmorethan10_status = true if ((@create_date + 10) < end_date) && (@close_date.nil? || (@close_date > end_date))
	end

	def generate_case_statistics(start_date, end_date, days_inactive)
		check_inactive(start_date, end_date, days_inactive) if ((@create_date <= end_date) && (@close_date.nil? || (@close_date > end_date)))
		check_open(start_date, end_date)
		check_closed(start_date, end_date)
		check_created(start_date, end_date)
		check_openmorethan10(start_date, end_date)
	end

end

Struct.new("Correspondence", :entry_date, :comment)