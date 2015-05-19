def GetWord(input, _idx, split_char) :
	x = input[_idx:].lstrip()
	
	_idx = len(input) - len(x)
	if len(x) == 0 :
		return "",len(input)
		
	_start = 0
	_end = _start
	
	if x[_start] == '{' :
		_depth = 1
		_start = _start + 1
		_end = _start
		while _end < len(x) :
			#print "_depth:", _depth
			if x[_end] == '{' :
				_depth += 1
			elif x[_end] == '}' :
				_depth -= 1
			if _depth == 0 :
				_idx = _idx + _end
				while _idx < len(input) :
				  #print _start, _end, _idx
					if input[_idx] == split_char :
						_idx += 1
						return x[_start:_end],_idx
					_idx += 1
				#print _start, _end, _idx
				return x[_start:_end],_idx
			_end += 1	
			
	while _end < len(x) :
		#print _end
		if x[_end] == split_char :
			_idx = _idx + _end + 1
			return x[_start:_end],_idx
		_end += 1			
	
	return x[_start:len(x)], len(input)	
	
			
def SplitEx(input, split_char) :
	x = []
	idx = 0
	while idx < len(input) :
		w = GetWord(input, idx, split_char)
		x.append(w[0])
		idx = w[1]
	return x		
		
def GetElemCount(array, elem) :
	print "GetElemCount:", array, elem
	count = 0
	for i in range(len(array)):
		if elem == array[i]:
			count+=1
	return count

def LastIndexOf(ss, c) :
	length = len(ss)
	for i in range(length) :
		if ss[length - 1 - i] == c :
			return length - 1 - i
	return -1
	
def RemovePostfix(ss) :
	i = -1
	localstr = ss
	while True :
		try:	
			j = localstr.index('.')
			i = i + j + 1
			localstr = localstr[j+1:]
		except:
			break
	if i >= 0 :
		return ss[0:i]
	return ss

def GetName(filename) :
	i1 = LastIndexOf(filename, '/')
	#print str(i1)
	i2 = LastIndexOf(filename, '\\')
	#print str(i2)
	if i1 < i2:
		i1 = i2
	if i1 >= 0 :
		ss = filename[i1+1:]
	else :
		ss = filename

	i1 = LastIndexOf(ss, '.')
	if i1 >= 0 :
		ss = ss[:i1]
	return ss

def GetDir(filename) :
	i1 = LastIndexOf(filename, '/')
	#print str(i1)
	i2 = LastIndexOf(filename, '\\')
	#print str(i2)
	if i1 < i2:
		i1 = i2
	if i1 >= 0 :
		ss = filename[0:i1]
	else :
		ss = "."
	return ss

def JoinPath(dirname, filename) :
        if dirname.endswith("/") or dirname.endswith("\\") :
                return dirname + filename
        else:
                return dirname + "/" + filename
                