from fuzzywuzzy import process,fuzz

def compare2list(leftl,rightl,resultdict):
    '''
    2 list, using right one to match left one 
    '''
    #这里如果用dict则不能重复，需要解决
    for item in range(len(leftl)):
        resultdict[item]=[leftl[item],process.extractOne(leftl[item], rightl,scorer=fuzz.token_sort_ratio)]