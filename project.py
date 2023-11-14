import json
import os
from openpyxl import Workbook
from openpyxl.styles import Font

num_of_parts = 5 # min = 1, max = 5

FILE_1 = 'data/rcv1-v2.topics.qrels.txt'
FILE_2 = [f'data/lyrl2004_vectors_test_pt{i}.dat.txt' for i in range(num_of_parts)]
FILE_3 = 'data/stem.termid.idf.map.txt'

JSON_1 = 'json_files/cat_dict.json'
JSON_2 = 'json_files/term_dict.json'
JSON_3 = 'json_files/stem_terms.json'
JSON_COMMAND_1 = 'json_files/command_1.json'
JSON_COMMAND_2 = 'json_files/command_2.json'
JSON_COMMAND_3 = 'json_files/command_3.json'
JSON_COMMAND_5 = 'json_files/command_5.json'

MAX_ROWS_XLSX = 1048570

"""
	Saves a collection as a json file in the given path
	Removes the file if it already exists
"""
def save_coll_as_json(collection, json_path):
	if os.path.exists(json_path):
		os.remove(json_path)

	with open(json_path, 'w') as json_file:
		json.dump(collection, json_file, indent=4)

"""
	We get a dictionary with categories as keys and 
	list of documents as values, like this:

	{'C11': ['2290'],
	 'C12': ['2295', '2299', '2305', '2306', '2311'],
	 'C13': ['2296', '2304'],
	 ...}

	We load num lines (n >= 0) else all of them (n < 0)
"""
def load_document_categories(file_path, num=-1):
	dc_dict = {}
	with open(file_path, "r", encoding="utf8") as dc:
		lines = dc.readlines()[:num] if num >= 0 else dc.readlines()
		for line in lines:
			pair = line[:-3] # throwing the 1's
			pair = pair.split(' ')

			if (len(pair) != 2):
				raise ValueError
			
			category = pair[0]
			document = pair[1]

			if category not in dc_dict.keys():
				dc_dict[category] = []

			dc_dict[category].append(document)		

	return dc_dict

"""
	We get a dictionary with terms as keys and 
	list of documents as values, like this:

	{"36975": [
		"254225",
		"254237",
		"280007",
		"280011",
		...]
	...}

	We load num lines (n >= 0) else all of them (n < 0)
"""
def load_document_terms(file_path, num=-1):
	dt_dict = {}
	for part in file_path:
		with open(part, "r", encoding="utf8") as dt_part:
			lines = dt_part.readlines()[:num] if num >= 0 else dt_part.readlines()
			for line in lines:
				temp = line.split(" ")
				document = temp[0]
				terms = [x.split(":")[0] for x in temp[2:]]
				
				for term in terms:
					if term not in dt_dict.keys():
						dt_dict[term] = []
				
				dt_dict[term].append(document)	

	return dt_dict

"""
	Every stem in the list has term_id = list:index + 1
"""
def load_terms_mapping(file_path):
	stems_terms = []
	with open(file_path, "r", encoding="utf8") as t:
		for line in t.readlines():
			stems_terms.append(line.split(" ")[:2][0])

	return stems_terms

"""
	Returns the score as float with five decimals
								|intersection|
	J(T, C) = ------------------------------------------------------
				|docs_per_category| + |docs_per_term| - |intersection|
"""
def jaccard_index(a, b, i):
	return "{:.5f}".format(i / (a + b - i))

def command_1(category, k, cat_dict, term_dict, stem_list):
	document_list = cat_dict[category]
	a = len(document_list)
	terms_score = []

	for term in term_dict.keys():
		docs_with_term = term_dict[term]
		b = len(docs_with_term)
		if (b > 0):
			intersection = len(set(document_list).intersection(docs_with_term))
			terms_score.append([stem_list[int(term)-1], jaccard_index(a, b, intersection)])

	terms_score.sort(key=lambda x: x[1])
	return terms_score[-k:]

def command_2(stem, k, cat_dict, term_dict, stem_list):
	term_id = stem_list.index(stem) + 1
	document_list = term_dict[str(term_id)]

	a = len(document_list)
	terms_score = []

	for category in cat_dict.keys():
		docs_of_category = cat_dict[category]
		b = len(docs_of_category)

		intersection = len(set(document_list).intersection(docs_of_category))
		terms_score.append([category, jaccard_index(a, b, intersection)])

	terms_score.sort(key=lambda x: x[1])
	return terms_score[-k:]

def command_3(stem, category, cat_dict, term_dict, stem_list):
	term_id = stem_list.index(stem) + 1
	document_list1 = cat_dict[category]
	document_list2 = term_dict[str(term_id)]

	a = len(document_list1)
	b = len(document_list2)

	intersection = len(set(document_list1).intersection(document_list2))

	return [category, jaccard_index(a, b, intersection)]

def command_4(filename, cat_dict, term_dict, stem_list):
	elem_list = []
	for category in cat_dict.keys():
		document_list = cat_dict[category]
		a = len(document_list)
		for term in term_dict.keys():
			docs_with_term = term_dict[term]
			b = len(docs_with_term)
			intersection = len(set(document_list).intersection(docs_with_term))
			
			elem = {
				"Stem": stem_list[int(term)-1],
				"Category": category,
				"JI Score": jaccard_index(a, b, intersection)
				}
			
			elem_list.append(elem)

	type_of_file = filename.split(".")[1]
	if (type_of_file == 'xlsx'):
		excel = Workbook()
		excel_sheet = excel.active
		excel_sheet.append(["Stem", "Category", "JI Score"])
		for index in ['A1', 'B1', 'C1']:
			excel_sheet[index].font = Font(size=14, bold=True)
		for i, elem in enumerate(elem_list):
			if (i >= MAX_ROWS_XLSX):
				break
			excel_sheet.append(list(elem.values()))
		if os.path.exists(filename):
			os.remove(filename)
		excel.save(filename)
	elif (type_of_file == 'json'):
		save_coll_as_json(elem_list, filename)

def command_5(doc_id, identifier, cat_dict, term_dict, stem_list):
	result = []
	if (identifier == '-c'):
		for category in cat_dict.keys():
			if doc_id in cat_dict[category]:
				result.append(category)
	elif (identifier == '-t'):
		for term in term_dict.keys():
			if doc_id in term_dict[term]:
				result.append(stem_list[int(term)-1])
	else:
		raise ValueError
	return result

def command_6(doc_id, identifier, cat_dict, term_dict, stem_list):
	return len(command_5(doc_id, identifier, cat_dict, term_dict, stem_list))
	
def cli(command, cat_dict, term_dict, stem_list):
	command_split = command.split(" ")
	command_symbol = command_split[0]

	if (command_symbol == '@'):
		category = command_split[1]
		k = int(command_split[2])
		return command_1(category, k, cat_dict, term_dict, stem_list)
	elif (command_symbol == '#'):
		stem = command_split[1]
		k = int(command_split[2])
		return command_2(stem, k, cat_dict, term_dict, stem_list)
	elif (command_symbol == '$'):
		stem = command_split[1]
		category = command_split[2]
		return command_3(stem, category, cat_dict, term_dict, stem_list)
	elif (command_symbol == '*'):
		filename = command_split[1]
		command_4(filename, cat_dict, term_dict, stem_list)
	elif (command_symbol == 'P'):
		doc_id = command_split[1]
		identifier = command_split[2]
		return command_5(doc_id, identifier, cat_dict, term_dict, stem_list)
	elif (command_symbol == 'C'):
		doc_id = command_split[1]
		identifier = command_split[2]
		return command_6(doc_id, identifier, cat_dict, term_dict, stem_list)
	else:
		raise ValueError

def main():
	cat_dict = load_document_categories(FILE_1)
	save_coll_as_json(cat_dict, JSON_1)

	term_dict = load_document_terms(FILE_2)
	save_coll_as_json(term_dict, JSON_2)

	stem_list = load_terms_mapping(FILE_3)
	save_coll_as_json(stem_list, JSON_3)

	# command_1_test = cli("@ E14 50", cat_dict, term_dict, stem_list)
	# save_coll_as_json(command_1_test, JSON_COMMAND_1)

	# command_2_test = cli("# winnie 50", cat_dict, term_dict, stem_list)
	# save_coll_as_json(command_2_test, JSON_COMMAND_2)

	# command_3_test = cli("$ winnie GWEA", cat_dict, term_dict, stem_list)
	# save_coll_as_json(command_3_test, JSON_COMMAND_3)

	# cli("* json_files/hello.xlsx", cat_dict, term_dict, stem_list)

	# command_5_test = cli("P 427211 -t", cat_dict, term_dict, stem_list)
	# save_coll_as_json(command_5_test, JSON_COMMAND_5)
	
	# command_6_test = cli("C 427211 -c", cat_dict, term_dict, stem_list)
	# print(command_6_test)

if (__name__ == "__main__"):
	main()
