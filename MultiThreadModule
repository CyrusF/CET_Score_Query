import threading
import queue
import time

def work():
    if not free.empty():
        x=free.get()
        global i
        i += 1
        print(">>> Get %s for processing %d"%(x,i))
        
        time.sleep(0.1)
        #Here put the query
        
        print("Free %s"%(x))
        free.put(x)

free=queue.LifoQueue(10)
for i in range(10):
    free.put("thread"+str(i+1))

i=0

while i<=100:
    t=threading.Thread(target=work)
    t.start()
    print("[%d]"%(i))#mean it is still trapped
    #t.join()
