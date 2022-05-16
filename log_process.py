import time
import datetime

def log(message,dateTimeFormat='%Y-%m-%d %H:%M:%S'):
    print('[{}] : {}'.format(datetime.datetime.now().strftime(dateTimeFormat),message))

def logTime(func,*args,**kwargs):
    '''
    This function is Decorator function
    '''
    def wrapperFunction(*args,**kwargs):
        functionName = func.__name__

        #start time
        startTime = time.time()

        #Show log start time before call function
        #log('Start function : {} with args {} kwargs {}'.format(functionName,args,kwargs))
        log('Start function : {}'.format(functionName))

        #Call function
        returnFunction = func(*args, **kwargs)

        #End time
        endTime = time.time()
        totalTime = endTime - startTime

        #Show log end time
        log('End function : {} - Total time : {}'.format(functionName, totalTime))

        return returnFunction
    return wrapperFunction
