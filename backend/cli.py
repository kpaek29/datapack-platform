#!/usr/bin/env python3
"""
DataPack Platform - Admin CLI
Manage users without exposing registration publicly
"""
import argparse
import sys
from .auth import create_user, get_users_db, save_users_db, get_password_hash

def add_user(username: str, password: str, email: str = None, name: str = None):
    """Add a new user"""
    try:
        user = create_user(username, password, email, name)
        print(f"✓ Created user: {username}")
        return True
    except ValueError as e:
        print(f"✗ Error: {e}")
        return False

def list_users():
    """List all users"""
    users = get_users_db()
    if not users:
        print("No users found")
        return
    
    print(f"\n{'Username':<20} {'Email':<30} {'Name':<20}")
    print("-" * 70)
    for username, data in users.items():
        email = data.get('email') or '-'
        name = data.get('full_name') or '-'
        print(f"{username:<20} {email:<30} {name:<20}")
    print()

def delete_user(username: str):
    """Delete a user"""
    users = get_users_db()
    if username not in users:
        print(f"✗ User '{username}' not found")
        return False
    
    del users[username]
    save_users_db(users)
    print(f"✓ Deleted user: {username}")
    return True

def reset_password(username: str, new_password: str):
    """Reset a user's password"""
    users = get_users_db()
    if username not in users:
        print(f"✗ User '{username}' not found")
        return False
    
    users[username]['hashed_password'] = get_password_hash(new_password)
    save_users_db(users)
    print(f"✓ Password reset for: {username}")
    return True

def main():
    parser = argparse.ArgumentParser(description="DataPack Platform Admin CLI")
    subparsers = parser.add_subparsers(dest='command', help='Commands')
    
    # add-user
    add_parser = subparsers.add_parser('add-user', help='Add a new user')
    add_parser.add_argument('username', help='Username')
    add_parser.add_argument('password', help='Password')
    add_parser.add_argument('--email', '-e', help='Email address')
    add_parser.add_argument('--name', '-n', help='Full name')
    
    # list-users
    subparsers.add_parser('list-users', help='List all users')
    
    # delete-user
    del_parser = subparsers.add_parser('delete-user', help='Delete a user')
    del_parser.add_argument('username', help='Username to delete')
    
    # reset-password
    reset_parser = subparsers.add_parser('reset-password', help='Reset user password')
    reset_parser.add_argument('username', help='Username')
    reset_parser.add_argument('password', help='New password')
    
    args = parser.parse_args()
    
    if args.command == 'add-user':
        success = add_user(args.username, args.password, args.email, args.name)
        sys.exit(0 if success else 1)
    elif args.command == 'list-users':
        list_users()
    elif args.command == 'delete-user':
        success = delete_user(args.username)
        sys.exit(0 if success else 1)
    elif args.command == 'reset-password':
        success = reset_password(args.username, args.password)
        sys.exit(0 if success else 1)
    else:
        parser.print_help()

if __name__ == '__main__':
    main()
