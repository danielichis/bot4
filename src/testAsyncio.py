import asyncio
import time

async def function1(i):
    print("running function ",i)
    await asyncio.sleep(i)
    print("finished function ",i)
async def function2():
    print("Hello ...")
    await asyncio.sleep(2)
    print("... World!")

async def main():
    tasks=[]
    for i in range(4):
        task=asyncio.create_task(function1(i+1))
        tasks.append(task)
    await asyncio.gather(*tasks)

asyncio.run(main())