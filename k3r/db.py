# -*- coding: utf-8 -*-
import pypyodbc


__author__ = "Виноградов А.Г. г.Белгород  август 2015"


class DB:
    def __init__(self, path_db):
        self.path_db = path_db

    def open(self):
        self.conn = pypyodbc.win_connect_mdb(self.path_db)
        self.cur = self.conn.cursor()

    def close(self):
        """Закрываем соединение"""
        self.cur.close()
        self.conn.close()

    def rs(self, sql):
        """Обработка запроса"""
        self.cur.execute(sql)
        rows = self.cur.fetchall()
        return rows
