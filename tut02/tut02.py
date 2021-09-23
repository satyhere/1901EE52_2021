def get_memory_score():
    rdef memory_score(input_nums) :
    memory=[]
    ans=0
        for i in input_nums :
            if len(memory) ==0:
            memory.append(i)
            else :
                present =0
            for j in memory :
                if j==i :
                    present=1
            if present :
                ans+=1
            else :
                memory.append(i)
                if len(memory)==6 :
                    memory.pop(0)
    return ans

    input_nums=[3,4,3,0,7,4,5,2,1,3]
    f=1
    for i in input_nums :
        if type(i)!=int :
            f=0

    if f==0 :
        print("The input list is not valid")
    else :
        print(memory_score(input_nums))