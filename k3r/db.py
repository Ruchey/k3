# -*- coding: utf-8 -*-
__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


import pypyodbc
import os


class DB:
    
    def __init__(self):
        pass

    def open(self, pathDB):
        '''Соединение с базой данных'''

        self.conn = pypyodbc.win_connect_mdb(pathDB)
        self.cur = self.conn.cursor()

    def close(self):
        '''Закрываем соединение'''
        
        if not self.conn:
            return None
        self.cur.close()
        self.conn.close()

    def rs(self, sql):
        """Обработка запроса"""
        
        try:
            self.cur.execute(sql)
            rows = self.cur.fetchall()
            return rows
        except:
            return []
