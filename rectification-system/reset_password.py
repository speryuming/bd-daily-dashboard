#!/usr/bin/env python3
"""Local account recovery for the rectification tracker."""

from __future__ import annotations

import getpass
import secrets
import sqlite3

import app


def main() -> None:
    app.init_db()
    with app.connect() as db:
        accounts = db.execute(
            "SELECT username,display_name,role FROM users WHERE active=1 ORDER BY id"
        ).fetchall()
        if not accounts:
            print("系统中没有可重置的账号。")
            return
        print("现有账号：")
        for row in accounts:
            role = "管理人员" if row["role"] == "manager" else "市场人员"
            print(f"  {row['username']}  {row['display_name']}  {role}")
        default = accounts[0]["username"] if len(accounts) == 1 else ""
        prompt = f"要重置的账号（直接回车使用 {default}）：" if default else "要重置的账号："
        username = input(prompt).strip() or default
        row = db.execute("SELECT * FROM users WHERE username=? AND active=1", (username,)).fetchone()
        if not row:
            print("账号不存在或已停用。")
            return
        new_username = input(f"新的登录账号（直接回车保持 {username}）：").strip() or username
        if len(new_username) < 2 or len(new_username) > 40:
            print("账号需为2至40个字符。")
            return
        if new_username != username and db.execute(
            "SELECT 1 FROM users WHERE username=?", (new_username,)
        ).fetchone():
            print("新的登录账号已存在。")
            return
        password = getpass.getpass("输入新密码（至少10个字符，输入时不会显示）：")
        confirm = getpass.getpass("再次输入新密码：")
        if password != confirm:
            print("两次密码不一致，未修改。")
            return
        if len(password) < 10:
            print("密码至少需要10个字符，未修改。")
            return
        salt = secrets.token_bytes(16)
        display_name = new_username if row["display_name"] == username else row["display_name"]
        try:
            db.execute("""UPDATE users SET username=?,display_name=?,salt=?,password_hash=?
                WHERE id=?""", (new_username, display_name, salt.hex(),
                app.password_hash(password, salt), row["id"]))
            db.execute("DELETE FROM sessions WHERE user_id=?", (row["id"],))
            db.execute("""UPDATE managers SET username=?,salt=?,password_hash=?
                WHERE username=?""", (new_username, salt.hex(),
                app.password_hash(password, salt), username))
        except sqlite3.IntegrityError:
            print("新的登录账号已存在，未修改。")
            return
    print(f"账号 {new_username} 的登录信息已更新，请返回网页重新登录。")


if __name__ == "__main__":
    main()
