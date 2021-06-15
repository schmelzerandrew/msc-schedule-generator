import random as r, openpyxl as op, os



CONFIG_TABLE_ROW_START = 9
CONFIG_TABLE_COLUMN_START = 2
CONFIG_TABLE_ROW_END = 22
CONFIG_TABLE_COLUMN_END = 8
TUTOR_NAME_CELL = (3,3)
TUTOR_WORK_AWARD_CELL = (3,8)
TUTOR_DESIRED_HOURS_CELL = (5,8)

MSC_WORKERS_NEEDED_FILENAME = 'MSC Hours of Operation.xlsx'
AVAILABILITY_FOLDER_NAME = 'student availability'

MSC_CONSTRAINTS_FILENAME = 'MSC Tutor Constraints.xlsx'
MSC_TUTOR_SCHEDULE_FILENAME = 'MSC Tutor Schedule.xlsx'

MAX_AWARD = 15

class ScheduleManager:

    def __init__(self):
        self.loaded = False
        self.worker_capacity = dict() #person: (preferred, allotted, open)
        self.worker_constraints = dict()  #time,day  : dict(name, pref)
        self.shifts = dict()              #time,day  : (CSB workers needed, SJU workers needed)
        self.total_available_hours = 0

    def initialize(self):
        """
        Checks if all the dependent files and folders are set up and ready. 
        """

        # if we cannot find the hours of operation file, that's a problem
        # if we cannot find the folder with the availability, that's a problem.
        # if we cannot find the folder to put the used availabilities, that's a problem

        curdir = os.getcwd()

        contents = os.listdir()

        all_good = True

        if not MSC_WORKERS_NEEDED_FILENAME in contents:
            print(f"'{MSC_WORKERS_NEEDED_FILENAME}' is a dependency of this program, please place that file into the same folder as this script: {os.getcwd()}")
            all_good = False
        if not AVAILABILITY_FOLDER_NAME in contents:
            print(f"'{AVAILABILITY_FOLDER_NAME}' folder is a dependency of this program. I've created that folder, place tutor availability forms into that folder.")
            os.mkdir(AVAILABILITY_FOLDER_NAME)
            all_good = False

        if len(os.listdir(AVAILABILITY_FOLDER_NAME)) < 5:
            file_count = os.listdir(AVAILABILITY_FOLDER_NAME)
            print(f"Not enough files found in the folder {AVAILABILITY_FOLDER_NAME} to continue executing. \n Found: {file_count} \t Minimum: 5.")
            all_good = False

        return all_good
        

    def load_msc_schedule(self):
        """
        loads the hours the MSC is open and how many workers are needed at each time
        from the configuration spreadsheet.
        """
        wb = op.open(r"MSC Hours of Operation.xlsx")
        ws = wb.active

        totalAvailableHours = 0

        #table starts at (9,2)
        #table ends at (22,8)
        for hour in range(CONFIG_TABLE_ROW_START,CONFIG_TABLE_ROW_END + 1):
            for day in range(CONFIG_TABLE_COLUMN_START, CONFIG_TABLE_COLUMN_END +1):
                try:
                    cell = ws.cell(hour,day)

                    workers_needed = self.parse_configuration_cell(cell)  # (CSB, SJU)
                    
                    totalAvailableHours += sum(workers_needed)
                    self.shifts[(hour,day)] = workers_needed
                    self.worker_constraints.setdefault((hour,day), dict())
                except Exception as e:
                    raise ValueError(f"Error found in spreadsheet at cell {(hour,day)}:", e)

        self.total_available_hours = totalAvailableHours
                
        return True

    def parse_configuration_cell(self, cell):
        v = cell.value
        priority = 0
        if v == None:
            return None
        if type(v) == int: #given no labels, default to CSB
            return (v,0)
        if type(v) == float:
            return (int(v),0)#same here, but floor to an int
        v = v.lower()
        if "," in v:
            parts = v.split(",")
            if "j" in parts[0].lower():
                sjuhours = parts[0].lower().strip("sju ")
                csbhours = parts[1].lower().strip("csb ")
                return (int(csbhours),int(sjuhours))
            elif "b" in parts[0].lower():
                sjuhours = parts[1].lower().strip("sju ")
                csbhours = parts[0].lower().strip("csb ")
                return (int(csbhours),int(sjuhours))
        elif "j" in v:
            sjuhours = v.lower().strip("sju ")
            csbhours = 0
            return (int(csbhours),int(sjuhours))
        elif "b" in v:
            sjuhours = 0
            csbhours = v.lower().strip("csb ")
            return (int(csbhours),int(sjuhours))
        elif v.isdecimal():
            sjuhours = 0
            csbhours = v
            return (int(csbhours),int(sjuhours))
        else:
            raise ValueError("Configuration spreadsheet has improper values within the table: recieved a non-integer.")
            

    def import_worker_schedules(self):
        """
        Reads the worker availability forms and loads them into the schedule constraints.
        Loops through the specified directory for .xlsx files, and sends them to parse.
        """

        files = os.listdir(AVAILABILITY_FOLDER_NAME)
        all_good = True
        for fn in files:
            if ".xlsx" in fn:
                try:
                    self.parse_availability_form(AVAILABILITY_FOLDER_NAME + "\\" + fn)
                except ValueError as ve:
                    all_good = False
                    print("Something's wrong in this file: \n" + str(ve))
                    
            else:
                print(f"WARNING: Non-spreadsheet file found in {AVAILABILITY_FOLDER_NAME}:  {fn}")
                

        return all_good

    
    def parse_availability_form(self, fn):
        """
        Parses an availability form and loads the data into the schedule constraints.
        """

        wb = op.open(fn)
        ws = wb.active

        worker_name = ws.cell(TUTOR_NAME_CELL[0],TUTOR_NAME_CELL[1]).value
        worker_name = str(worker_name).strip()

        if worker_name in ("", "None"):
            raise ValueError("Worker name field left blank.")

        award_hours = ws.cell(TUTOR_WORK_AWARD_CELL[0],TUTOR_WORK_AWARD_CELL[1]).value
        award_hours = int(award_hours)

        if (award_hours <= 0): 
            raise ValueError(f"Improper number of work award hours: {award_hours}.")
        if award_hours > MAX_AWARD:
            raise ValueError(f"Improper number of work award hours: maximum exceeded: {MAX_AWARD}.")
        
        desired_hours = ws.cell(TUTOR_DESIRED_HOURS_CELL[0],TUTOR_DESIRED_HOURS_CELL[1]).value
        desired_hours = int(desired_hours)

        if desired_hours < 0:
            raise ValueError(f"Improper number of desired hours: {desired_hours}.")
        if desired_hours > award_hours:
            raise ValueError(f"Improper number of desired hours; exceeds work award: {desired_hours}.")

        total_open_hours = 0
        error_spots = []
        for hour in range(CONFIG_TABLE_ROW_START,CONFIG_TABLE_ROW_END + 1):
            for day in range(CONFIG_TABLE_COLUMN_START, CONFIG_TABLE_COLUMN_END +1):
                cell = ws.cell(hour,day)
                
                preference = self.parse_worker_preference(cell) #(CSB, SJU)

                if type(preference) == tuple:
                    total_open_hours += 1

                    constraint_slot = self.worker_constraints[(hour,day)]
                    
                    constraint_slot[worker_name] = preference

                    self.worker_constraints[(hour,day)] = constraint_slot
                elif preference == "e":
                    error_spots.append(cell.coordinate)
                #elif preference == None do nothing


        if len(error_spots) > 0:
            raise ValueError("Improper entries in the following cells: \n" + "\n".join(error_spots))
                    
        wb.close()

        self.worker_capacity[worker_name] = (desired_hours,award_hours, total_open_hours) 
        return 

    def parse_worker_preference(self, cell):
        """
        Scans the cell and returns the preferences of the worker
        """
        v = cell.value
        if v == None:
            return None
        if type(v) == int: #given no labels, default to CSB
            return (v,-1)
        if type(v) == float:
            return (int(v),0)#same here, but floor to an int
    
        v = v.lower()
        if "x" in v:
            return None
        elif "-" in v:
            return None
        elif "csb or sju" in v:
            v = v.strip("csb or sju")
            return (int(v),int(v))
        elif "j" in v:
            sjupriority = v.strip("sju ")
            csbpriority = -1
            return (int(csbpriority),int(sjupriority))
        elif "b" in v:
            sjupriority = v.strip("csb ")
            csbpriority = -1
            return (int(csbpriority),int(sjupriority))
        elif v.isdecimal():
            sjuhours = -1
            csbpriority = v
            return (int(csbpriority),int(sjupriority))
        else:
            return "e"

    def create_default_schedule(self):
        """Takes into account the loaded schedule constraints and develops
        a schedule that satisfies them.

        Uses randomness to reduce inherent bias from hash or
        alphabetical considerations. Non-deterministic.
        """

        ps = PotentialSchedule(self.worker_capacity, self.shifts)
        
        IsFirstPriority  = lambda x,y: x[1][y] == 1
        IsSecondPriority = lambda x,y: x[1][y] == 2
        IsThirdPriority  = lambda x,y: x[1][y] == 0
        

        priorityLevels = (IsFirstPriority, IsSecondPriority, IsThirdPriority)


        for campus in (0,1):
            for timeslot in self.shifts.keys():
                
                workingAtOtherCampus = lambda x: x in ps.schedule[(campus+1)%2].setdefault(timeslot,[])
                
                constraints = self.worker_constraints[timeslot]
                workers = constraints.items() #(name, (CSB,SJU))
                
                workers_needed = self.shifts[timeslot][campus]

                scheduled_workers = []
                

                num_yet_needed = workers_needed

                # run through for people with more hours desired
                for priority in priorityLevels:

                    filter_fxn = lambda x: priority(x,campus) and ps.desires_more_hours(x) and not workingAtOtherCampus(x)
                    
                    workers_with_desire = list(filter(filter_fxn, workers))


                    if len(workers_with_desire) < num_yet_needed:
                        scheduled_workers.extend(workers_with_desire)
                    else:
                        scheduled_workers.extend(r.sample(workers_with_desire, k = num_yet_needed))
                        num_yet_needed = workers_needed - len (scheduled_workers)
                        break
                    
                    num_yet_needed = workers_needed - len (scheduled_workers)

                


                # if nobody WANTS to work but we still need it filled
                if num_yet_needed > 0:
                    # run through for people with more hours allotted

                    # don't allow people to be scheduled twice!
                    
                    for priority in priorityLevels:
                        filter_fxn = lambda x: priority(x, campus) and ps.allotted_more_hours(x)
                        filter_fxn2 = lambda x: filter_fxn(x) and x not in scheduled_workers and not workingAtOtherCampus(x)
                    
                        workers_with_allotment = list(filter(filter_fxn2, workers))


                        if len(workers_with_allotment) < num_yet_needed:
                            scheduled_workers.extend(workers_with_allotment)
                        else:
                            scheduled_workers.extend(r.sample(workers_with_allotment, k = num_yet_needed))
                            break
                        
                        num_yet_needed = workers_needed - len (scheduled_workers)
                    
                # there's a chance we still need more people but we just don't have them.

                ps.add_workers_to_slot(timeslot, campus, scheduled_workers)

        return ps


    def successor(self, ps, n_changes):

        child = PotentialSchedule(ps.worker_capacity, ps.shifts, ps.worker_slotted_hrs, ps.schedule)




        for campus in (0,1):

            not_negative =      lambda x: x[1][campus] != -1
            has_more_hours =    lambda x: child.allotted_more_hours(x)
            positive_priority = lambda x: x[1][campus] > 0
            
            nonzero_shifts = list(filter(lambda x: ps.shifts[x][campus], ps.shifts.keys()))
            
            reroll_shifts = r.sample(nonzero_shifts, n_changes)
            for timeslot in reroll_shifts:
                
                for worker in ps.schedule[campus][timeslot]:

                    child.worker_slotted_hrs[worker[0]] -= 1

                child.schedule[campus][timeslot] = []

            for timeslot in reroll_shifts:
                
                
                available_workers = list(filter(lambda x: not_negative(x) and has_more_hours(x), self.worker_constraints[timeslot].items()))

                worker_count = min(len(available_workers), ps.shifts[timeslot][campus])
                
                child.add_workers_to_slot(timeslot, campus, r.sample(available_workers, worker_count))

        return child


    def write_schedule_to_spreadsheet(self, ps):

        wb = op.open(MSC_TUTOR_SCHEDULE_FILENAME)
        ws = wb.active
        for campus in (0,1):
            ws = wb[["CSB","SJU"][campus]]
            for timeslot, workers in ps.schedule[campus].items():
                if len(workers) == 0:
                    ws.cell(timeslot[0],timeslot[1]).value = "-"
                else:
                    ws.cell(timeslot[0],timeslot[1]).value = ",\n".join(x[0] for x in workers)

        wb.save(MSC_TUTOR_SCHEDULE_FILENAME)
        wb.close()
        
        pass

    def write_constraints_to_spreadsheet(self):
        wb = op.open(MSC_CONSTRAINTS_FILENAME)
        ws = wb.active
        for campus in (0,1):
            ws = wb[["CSB","SJU"][campus]]
            for timeslot, constraints in self.worker_constraints.items():
                if not self.shifts[timeslot][campus]:
                    ws.cell(timeslot[0],timeslot[1]).value = None
                    continue
                people = ""
                for worker, preference in constraints.items():
                    if preference[campus] > 0:
                        people += worker + f": {preference[campus]}\n "
                if people == "":
                    for worker, preference in constraints.items():
                        if preference[campus] == 0:
                            people += worker + f": {preference[campus]}\n "
                ws.cell(timeslot[0],timeslot[1]).value = people

        wb.save(MSC_CONSTRAINTS_FILENAME)
        ws = wb.active

class PotentialSchedule:

    def __init__(self, wc, shifts, wsh = None, s = None):
        self.worker_capacity = wc.copy() #person: (preferred, allotted, open)
        
        if wsh == None:
            self.worker_slotted_hrs = dict() #person: n
        else:
            self.worker_slotted_hrs = wsh.copy()

            
        self.shifts = shifts.copy()


        if s == None:
            self.schedule = [dict(),dict()] #CSB,SJU (hour,day) : [(person, preference)]
        else:
            self.schedule = [s[0].copy(),s[1].copy()]
        
            
    def desires_more_hours(self, worker):
        """
        Returns true if the scheduler has not yet filled this workers desired
        hours.
        """
        current_hrs = self.worker_slotted_hrs.setdefault(worker[0], 0)
        desired_hrs = self.worker_capacity[worker[0]][0]
        return current_hrs < desired_hrs


    def allotted_more_hours(self, worker):
        """
        Returns true if the scheduler has not yet filled this workers allotted
        hours.
        """
        current_hrs = self.worker_slotted_hrs.setdefault(worker[0], 0)
        allotted_hrs = self.worker_capacity[worker[0]][1]
        return current_hrs < allotted_hrs

    def add_workers_to_slot(self, timeslot, campus, workers):
        """
        Places workers into the schedule, and updates workers' hour counts.
        """
        for worker in workers:
            current_hrs = self.worker_slotted_hrs.setdefault(worker[0],0)
            self.worker_slotted_hrs[worker[0]] += 1

        current_workers = self.schedule[campus].setdefault(timeslot, [])
        current_workers.extend(workers)

    def count_gaps(self):
        gaps = 0
        for campus in (0,1):
            for timeslot in self.schedule[campus].keys():
                holes = self.shifts[timeslot][campus] - len(self.schedule[campus][timeslot])
                gaps += holes
        return gaps

    def geometric_mean_desired(self):
        
        accum = 1
        count = 0
        for worker, prefs in self.worker_capacity.items():
            count += 1

            proportion = self.worker_slotted_hrs[worker] / prefs[0]


            # no bonus points for working more than you want to.
            # going past is just as bad as going below.
            regulated = 1 - abs(proportion - 1)
            
            accum*= proportion

        return pow(accum, 1/count)

    def mean_desired_weighted(self):
        count = 0
        accum = 0
        for worker, prefs in self.worker_capacity.items():
            proportion = self.worker_slotted_hrs[worker] / prefs[0]
            
            if prefs[2]/prefs[0] > 2: #if you don't have many open hours, your desires aren't counted
                accum += proportion
                count += 1

        return accum /count

    def avg_priority(self):
        count = 0
        accum = 0
        for campus in (0,1):
            for shift in self.schedule[campus].values():
                for person in shift:
                    if person[1][campus] == 1:
                        accum += 2
                    elif person[1][campus] == 2:
                        accum += 1
                    count += 1

        return accum / count

    def min_hrs_filled(self):
        return min(self.worker_slotted_hrs.values())

    def min_hrs_proportion(self):
        return min(self.worker_slotted_hrs[worker] / prefs[0] for worker, prefs in self.worker_capacity.items())

    def avg_trips_in(self):
        n_times = 0
        for campus in (0,1):
            for timeslot in self.schedule[campus].keys():
                workers = self.schedule[campus][timeslot]
                for worker in workers:
                    if worker not in self.schedule[campus][(timeslot[0]-1,timeslot[1])]:
                        n_times += 1
        return n_times/ len(self.worker_capacity)


    def report_scores(self):
        out = ""
        out += f" Avg Priority: {self.avg_priority():.2f}\n"
        out += f" Total hours filled: {sum(self.worker_slotted_hrs.values())}\n"
        out += f" Min hrs filled: {self.min_hrs_filled()}\n"
        out += f" Min hrs proportion: {self.min_hrs_proportion()}\n"
        out += f" Mean desired (weighted) filled: {self.mean_desired_weighted():.2f}\n"
        out += f" Mean trips in: {self.avg_trips_in():.2f}\n"
        out += f" Geom mean desired: {self.geometric_mean_desired():.2f}\n"
        out += f" Gaps in schedule: {self.count_gaps():.2f}\n"
        out += f" Score: {-self.evaluate():.2f}\n\n"
        return out

    def report_workers(self):

        workers = list(self.worker_capacity.keys())
        
        proportion = lambda worker: self.worker_slotted_hrs[worker]/self.worker_capacity[worker][0]
        
        workers.sort(key = proportion, reverse = True)

        out = ""
        for worker in workers:
            prefs = self.worker_capacity[worker]#person: (preferred, allotted, open)
            scheduled = self.worker_slotted_hrs[worker]
            out += f"{worker} \n- Desired: {prefs[0]:>2} Scheduled: {scheduled:>2} Proportion: {proportion(worker):>4.2f} Open: {prefs[2]}\n"
            
        return out

    def report(self):
        print(self.report_scores())
        print(self.report_workers())

    def write_report(self):
        report = open("report.txt", "w")
        report.write(self.report_scores())
        report.write(self.report_workers())
        report.close()
        
        

    def evaluate(self):

        score = 0

        score += self.count_gaps() * -1000

        score += self.min_hrs_filled() * 100

        score += self.min_hrs_proportion() ** 2 * 100

        score += sum(self.worker_slotted_hrs.values())



        score /= self.avg_trips_in()     #[1,12+)

        score *= self.avg_priority() ** 2 + 1   #[0,2]

        score *= self.mean_desired_weighted() + 1 

        return -score
            
        
        
if __name__ == "__main__":
    sm = ScheduleManager()
    sm.initialize()        
    sm.load_msc_schedule()
    sm.import_worker_schedules()
    ps = sm.create_default_schedule()
    ps.report()

