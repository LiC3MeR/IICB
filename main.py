from pip install pywin32 import Dispatch as d
from fuzzywuzzy.fuzz import ratio as rat
import webbrowser as web
import speech_recognition as sr

class Assistant:
	def __init__(self):
		self.engine = d("Sapi.SpVoice")

		self.cmds = {
			('привет', 'здаров', 'здравствуй') : self.hello,
			('пока', 'свали', 'уйди', 'ушёл', 'не мешай') : self.bye
		}

		self.hellos = ['привет', 'здаров', 'здравствуй']

		self.dels = [
			'сара', 'пожалуйста', 'может', 'давай',
			'наверное'] + self.hellos

		self.keys = [
			'открой', 'включи', 'выключи', 'найди', 'зайди', 'запусти',
			'привет', 'здаров', 'здравствуй', 'пока', 'свали', 'уйди', 'ушёл', 'мешай'
		]

		self.m = sr.Microphone()
		self.r = sr.Recognizer()
		self.r.energy_threshold = 1000
		self.r.pause_threshold = 0.5


	def keys_count(self, task):
		count = 0
		for key in self.keys:
			count += task.count(key)
		return count


	def clear_cmd(self, text):

		save_task = text

		def check(word):
			flag = 0
			for i in self.keys:
				if i in word:
					flag = 1
					return 1
			if flag == 0:
				return 0

		if text:
			for word in self.dels:
				text = text.replace(word, '').replace('  ', ' ').strip()

		tasks = []

		for key in self.keys:
			if text:
				task = text.split()
				if key in task:
	
					var = [' '.join(task[0:task.index(key)]), ' '.join(task[task.index(key):])]
					for task in var:
						if not(task in tasks):
							if self.keys_count(task) < 2:
								tasks.append(task)

		if save_task:
			for i in self.hellos:
				if i in save_task:
					tasks.insert(0, i)
					break


		tasks = [x for x in tasks if x]
		return tasks


	def hello(self):
		self.say('Привет!')


	def bye(self):
		self.say('Ну пока!')
		exit()


	def opener(self, task):
		self.say('Сейчас открою')
		links = {
			('youtube', 'ютюб', 'ютуб') : 'https://www.youtube.com/',
			('вк', 'vk', 'контакт') : 'https://vk.com/feed'
		}
		for vals in links:
			for word in vals:
				if rat(word, task) > 75:
					web.open(links[vals])
					break


	def starter(self, task):
		dels = ('включи', 'выключи')
		if task.startswith('включи '):
			task = task.replace('включи ', '', 1).strip()
			print(f'Включаю {task}')
		elif task.startswith('выключи '):
			task = task.replace('выключи ', '', 1).strip()
			print(f'Выключаю {task}')


	def search(self, task):
		self.say('Сейчас мы всё найдём')
		for i in ('найди', 'найти'):
			task = task.replace(i, '').replace('  ', ' ').strip()
		print(f'Ищу {task}')
		web.open(f'http://www.google.com/search?btnG=1&q={task}')


	def listen(self):
		text = ''
		print('Я вас слушаю: ')
		with self.m as sourse:
			self.r.adjust_for_ambient_noise(sourse)
		while text == '':
			with self.m as sourse:
				audio = self.r.listen(sourse)
				try:
					text = (self.r.recognize_google(audio, language="ru-RU")).lower()
				except:
					pass
		else:
			return( text )


	def cmd_exe(self):
		tasks = self.clear_cmd(self.listen())
		print(tasks)

		for task in tasks:
			if task.startswith(('открой', 'зайди', 'запусти')):
				self.opener(' '.join(task.split()[1:]))
			elif task.startswith(('включи', 'выключи')):
				self.starter(task)
			elif task.startswith(('найди', 'найти')):
				self.search(task)
			else:
				print(task)
				for vals in self.cmds:
					for val in vals:
						if rat(val, task) > 80:
							print(vals)
							self.cmds[vals]()
							break


	def say(self, text):
		self.engine.Speak(text)

sara = Assistant()
while True:
	sara.cmd_exe()