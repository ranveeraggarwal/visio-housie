import os
import pdb
import random
import win32com.client

visio_handle = win32com.client.Dispatch("Visio.Application")

application = visio_handle.Application
documents = visio_handle.Documents

__location__ = os.path.realpath(
	os.path.join(os.getcwd(), os.path.dirname(__file__)))
current_document = documents.Open(
	os.path.join(__location__, 'Housie.vsdx'))

active_page = application.ActivePage

unrevealed = [i for i in range(1, 91)]
random.shuffle(unrevealed)

print('Welcome to Visio Housie!')

active_page.Shapes[3].Text = ''

while len(unrevealed):
	print('Press <Return> for new number or s <Return> for Shuffle ', end='')
	inp = input()
	if inp == 's':
		random.shuffle(unrevealed)
	else:
		current_number = unrevealed[0]
		unrevealed = unrevealed[1:]
		active_page.Shapes[0].Text = current_number
		if active_page.Shapes[3].Text == '':
			active_page.Shapes[3].Text = str(current_number)
		else: 
			active_page.Shapes[3].Text = active_page.Shapes[3].Text + ', ' + str(current_number)

print('Thank you for playing Visio Housie!')