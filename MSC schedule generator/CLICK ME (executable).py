import os, time, sys
from schedulemanager import *
import matplotlib.pyplot as plt
import random, math



NUM_CHANGES = 8

def main(sm, temp = 10000, coolingRate = 0.005):

    initTemp = temp
    
    trackEnergy = [] #this is only for graphing purposes later
    trackTemp =[] #this is only for graphing purposes later
    trackBest=[] #this is only for graphing purposes later

    
    ps = sm.create_default_schedule()

    print("Starting annealing process.")
    start_time = time.perf_counter()
    
    #calculate the energy
    energy = ps.evaluate()#for graphing
    trackEnergy.append(energy)#this is only for graphing purposes later
    trackBest.append(energy)#this is only for graphing purposes later
    trackTemp.append(temp)#this is only for graphing purposes later

    #print(f"The inital rating of this schedule = {energy}")
    

    #this is the best so far
    best_schedule = ps
    best_energy = energy
    
    #slowly "cool" the system
    while( temp >1):
        new_schedule = sm.successor(ps, NUM_CHANGES)
        
        #get current energy
        current_energy = ps.evaluate()
        new_energy = new_schedule.evaluate()
        
        #decide if should accept the neighborEnergy
        '''If it is not better than the current energy, then make it the current state
        with probability p as defined by protocal.  
        This step is usually implemented by invoking a random number generator to produce
        a number in the range of [0,1].  
        If that number is less than p, then the move is accepted.  
        Otherwise, do nothing.
        '''
        if(acceptanceProb(current_energy, new_energy, temp) > random.random()):
            ps = new_schedule
            
        #keep track of best so far
        if(new_energy < best_energy):
            best_schedule = new_schedule
            best_energy = new_energy
            
        trackEnergy.append(new_energy) #for graphing
        trackBest.append(best_energy) #for graphing
        trackTemp.append(temp) #for graphing
        #cool system
        temp *= 1-coolingRate

    end_time = time.perf_counter()

    t_diff = end_time - start_time


    print("Annealing completed.")
    print(f"Took {t_diff} seconds.")
    print(f"Best score was {-best_energy:.2f} points.")

    #make the graphs
    plotDistanceChanges(trackTemp, trackEnergy, "tracking energy over temp change")
    plotDistanceChanges(trackTemp, trackBest, "tracking best energy over temp change")
    
    return best_schedule, best_energy, t_diff
    
def acceptanceProb(energy, newEnergy, temperature):
    '''This calculation determines if we accept the new state or not
       If the new state is better - we always accept items - return 1
       If the new state is not better - we accept it based on a probability
    '''
    if(newEnergy< energy):
        #print("new is better")
        return 1.0
    return math.exp((energy - newEnergy)/temperature)

def plotDistanceChanges(dist, temp, title):
    '''
        This makes a plot to show distance changes while the
        temperature changes
    '''
    plt.title(title)
    plt.xlabel('temperature')
    plt.ylabel('distance')
    plt.plot(dist, temp, "ro-")
    plt.show()


def testingSuite():
    import itertools as it

    sm = ScheduleManager()
    try:
        assert sm.initialize()        
        assert sm.load_msc_schedule()
        assert sm.import_worker_schedules()
    except AssertionError as ae:
        print("System aborting...")
        time.sleep(10)
        return

    temps = [x*x * 1000 for x in range(1,11)]
    coolingRates = [0.001, 0.005,0.01,0.02]

    iterations = 10

    print(f"Launching {len(temps)*len(coolingRates)} experiments for {iterations} iterations each.")

    min_energy = float("inf")
    best_params = None
    start_time = time.time()
    for init_temp, cooling_rate in it.product(temps,coolingRates):
        accumTime = 0
        
        for i in range(iterations):
            
            best_schedule, best_energy, annealing_time = main(sm, init_temp, cooling_rate)
            
            if best_energy < min_energy:
                best_params = (init_temp, cooling_rate)
                min_energy = best_energy

            accumTime += annealing_time

        print(f"T: {init_temp} \t CR: {cooling_rate} \t Mean Run Time: {accumTime/iterations:.4f}")
    print("FINAL REPORT:")
    print(f"Params yielding the best score: {best_params}")
    print(f"Best score: {-min_energy:.2f}")
    end_time = time.time()
    t_diff = end_time - start_time

    print(f"We conducted {len(temps)*len(coolingRates)} experiments for {iterations} iterations each.")
    print(f"Experimentation took {t_diff:.2f} seconds.")


if __name__ == "__main__":

    sm = ScheduleManager()
    try:
        assert sm.initialize()        
        assert sm.load_msc_schedule()
        assert sm.import_worker_schedules()
    except AssertionError as ae:
        print("System aborting...")
        time.sleep(10)


    best_schedule, best_energy, annealing_time = main(sm)
    sm.write_schedule_to_spreadsheet(best_schedule)
    sm.write_constraints_to_spreadsheet()
    best_schedule.report()
    best_schedule.write_report()
    time.sleep(100)

    

