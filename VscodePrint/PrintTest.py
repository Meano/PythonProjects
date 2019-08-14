from multiprocessing import Pool, cpu_count
import traceback
import sys

def printfunc(msg):
    for i in range(20):
        print "TEST "+ str(msg) +" Count:" + str(i)
    return 0

if __name__ == "__main__":
    try:
        testlist = list(range(1000))
        pool_size = cpu_count()
        pool = Pool(processes=pool_size)
        print "Pool Size:", pool_size
        pool.map(printfunc, testlist)
        pool.close()
        pool.join()
    except Exception,e:
        pool.terminate()
        traceback.print_exc()
        sys.exit(1)