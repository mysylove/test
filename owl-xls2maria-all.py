# -*- coding: utf-8 -*-

import openpyxl
import mysql.connector
import os
import sys, traceback
import logging
import pymysql

result = []

# 엑셀 데이터를 tuple 형식으로 저장 학번이 201600000 일 경우 16만 저장하도록 파싱 
def excel_to_list(filename):
	wb = openpyxl.load_workbook(filename)
	ws = wb.active
	tmp_data = []
	idx = 0
 
	for r in ws.rows:
		if idx < 3:
			idx = idx + 1
			continue
		a0 = r[0].value
		b0 = r[1].value
		c0 = r[2].value
		d0 = r[3].value
		e0 = r[4].value
		f0 = r[5].value
		g0 = r[6].value
		h0 = r[7].value
		i0 = r[8].value
		j0 = r[9].value
		k0 = r[10].value
		l0 = r[11].value
		m0 = r[12].value
		n0 = r[13].value
		o0 = r[14].value
		p0 = r[15].value
		q0 = r[16].value
		r0 = r[17].value
		s0 = r[18].value
		t0 = r[19].value
		u0 = r[20].value
		v0 = r[21].value
		w0 = r[22].value
		x0 = r[23].value
		y0 = r[24].value
		z0 = r[25].value
		aa0 = r[26].value
		ab0 = r[27].value
		ac0 = r[28].value
		ad0 = r[29].value
		ae0 = r[30].value
		af0 = r[31].value
		ag0 = r[32].value
		ah0 = r[33].value
		ai0 = r[34].value
		aj0 = r[35].value
		ak0 = r[36].value
		al0 = r[37].value
		am0 = r[38].value
		an0 = r[39].value
		ao0 = r[40].value
		ap0 = r[41].value
		aq0 = r[42].value
		ar0 = r[43].value
		as0 = r[44].value
		at0 = r[45].value
		au0 = r[46].value
		av0 = r[47].value
		aw0 = r[48].value
		ax0 = r[49].value
		ay0 = r[50].value
		bl0 = r[63].value

		tmp_data.append(a0)
		tmp_data.append(b0)
		tmp_data.append(c0)
		tmp_data.append(d0)
		tmp_data.append(e0)
		tmp_data.append(f0)
		tmp_data.append(g0)
		tmp_data.append(h0)
		tmp_data.append(i0)
		tmp_data.append(j0)
		tmp_data.append(k0)
		tmp_data.append(l0)
		tmp_data.append(m0)
		tmp_data.append(n0)
		tmp_data.append(o0)
		tmp_data.append(p0)
		tmp_data.append(q0)
		tmp_data.append(r0)
		tmp_data.append(s0)
		tmp_data.append(t0)
		tmp_data.append(u0)
		tmp_data.append(v0)
		tmp_data.append(w0)
		tmp_data.append(x0)
		tmp_data.append(y0)
		tmp_data.append(z0)
		tmp_data.append(aa0)
		tmp_data.append(ab0)
		tmp_data.append(ac0)
		tmp_data.append(ad0)
		tmp_data.append(ae0)
		tmp_data.append(af0)
		tmp_data.append(ag0)
		tmp_data.append(ah0)
		tmp_data.append(ai0)
		tmp_data.append(aj0)
		tmp_data.append(ak0)
		tmp_data.append(al0)
		tmp_data.append(am0)
		tmp_data.append(an0)
		tmp_data.append(ao0)
		tmp_data.append(ap0)
		tmp_data.append(aq0)
		tmp_data.append(ar0)
		tmp_data.append(as0)
		tmp_data.append(at0)
		tmp_data.append(au0)
		tmp_data.append(av0)
		tmp_data.append(aw0)
		tmp_data.append(ax0)
		tmp_data.append(ay0)
		tmp_data.append(bl0)

		result.append(tuple(tmp_data))
		tmp_data = []

		#if name is not None and str(stuNum).isdigit():
		#	tmp_data.append(str(stuNum)[2:4])
		#	tmp_data.append(name)
		#if len(tmp_data) == 2:
		#	result.append(tuple(tmp_data))
		#	tmp_data = []

# mysql 테이블에 튜플 데이터 삽입
def mysql_insert(db,table,data):
	try:
		cursor = db.cursor()
		#sql = "INSERT INTO"+table+" (stuNum,name) VALUES (%s, %s)"
		#sql = "INSERT INTO "+table+" (a0, b0, c0, d0, e0, f0, g0, h0, i0, j0, k0, l0, m0, n0, o0, p0, q0, r0, s0, t0, u0, v0, w0, x0, y0, z0, aa0, ab0, ac0) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
		sql = "INSERT INTO "+table+" (a0, b0, c0, d0, e0, f0, g0, h0, i0, j0, k0, l0, m0, n0, o0, p0, q0, r0, s0, t0, u0, v0, w0, x0, y0, z0, aa0, ab0, ac0, ad0, ae0, af0, ag0, ah0, ai0, aj0, ak0, al0, am0, an0, ao0, ap0, aq0, ar0, as0, at0, au0, av0, aw0, ax0, ay0, bl0) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
		cursor.executemany(sql,data)
		db.commit()
		print("[+] Insertion success\n")
	except:
		logging.error(traceback.format_exc())
		print("[ERROR] Insertion failed\n")

# mysql 테이블의 기존 데이터를 삭제
def table_clear(db,table):
	try:
		cursor = db.cursor()
		cursor.execute("TRUNCATE TABLE "+table)
	except:
		print("[ERROR] Truncate failed\n")

# 파일이 존재하는지 체크
# return : True or False
def fileCheck(filename):
	return os.path.isfile("./"+filename)

def main():
	db = pymysql.connect(
		host="localhost",
		user="root",
		passwd="wowjddl!@34",
		db="test",
		charset='utf8'
	)
	print("* * * * * * OWL 엑셀 변환 데이터 삽입 자동화 프로그램* * * * * *")
	print("[*] exit 입력 시 종료됩니다.")
	print("[*] 엑셀 파일은 .py 파일과 같은 디렉토리에 존재해야 합니다.")
	print("[*] 테이블이 존재하지 않을 경우 생성 후 가능합니다.\n")
	while True:
		filename = input("1. 엑셀 파일명 입력 : ")
		if fileCheck(filename) is True:
			table = input("3. 데이터베이스 테이블명 입력 : ")
			excel_to_list(filename)
			answer = input("[*] 테이블 내 기존 데이터가 삭제됩니다. 진행하시겠습니까? (Y,n) : ")
			if answer == "Y":
				table_clear(db,table)
				mysql_insert(db,table,result)
				break
			else:
				continue
		elif filename == "exit":
			db.close()
			sys.exit(1)
		else:
			print("[ERROR] 파일이 존재하지 않습니다.")
			continue
		print("\n")

	sys.exit(1)

if __name__ == "__main__":
    main()