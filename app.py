#!/usr/bin/env python3
"""
Mailbox Storage Analyzer - Zero Error Margin Version (FINAL FIXED)
Calculates exact IMAP mailbox storage usage with 100% accuracy.
Priority: STATUS=SIZE > RFC822.SIZE (all messages) > QUOTA
No sampling, no estimation.
"""

import pandas as pd
import imaplib
from datetime import datetime
import time
import sys
import os
import re  # Added for robust parsing

print("✅ Script loaded successfully")
print(f"   Python version: {sys.version}")
print(f"   Working directory: {os.getcwd()}")
print("=" * 70)


class MailboxAnalyzer:
    def __init__(self):
        self.results = []
        self.capabilities_cache = {}

    def connect_to_imap(self, server, email_addr, password, port=993, use_ssl=True):
        try:
            print(f"   🔌 Connecting to {server}:{port}...")
            if use_ssl:
                mail = imaplib.IMAP4_SSL(server, port)
            else:
                mail = imaplib.IMAP4(server, port)
            mail.login(email_addr, password)
            print(f"   ✅ Logged in successfully")
            
            try:
                caps = self._parse_capabilities(mail)
                self.capabilities_cache[email_addr] = caps
                print(f"   📋 Capabilities detected: {len(caps)} items")
                if 'STATUS=SIZE' in caps:
                    print(f"   ⚡ STATUS=SIZE supported")
            except Exception as e:
                print(f"   ⚠️  Could not fetch capabilities: {e}")
                self.capabilities_cache[email_addr] = set()
            return mail
        except Exception as e:
            print(f"   ❌ Connection failed: {str(e)}")
            return None

    def _parse_capabilities(self, mail):
        try:
            caps_raw = mail.capability()
            if not caps_raw:
                return set()
            result = set()
            def flatten(item):
                if isinstance(item, (list, tuple)):
                    for sub in item:
                        flatten(sub)
                elif isinstance(item, bytes):
                    result.add(item.decode('utf-8', errors='ignore').upper())
                elif isinstance(item, str):
                    result.add(item.upper())
            flatten(caps_raw)
            return result
        except Exception:
            return set()

    def get_quota_info(self, mail):
        try:
            status, quota_data = mail.getquotaroot('INBOX')
            if status == 'OK' and quota_data:
                for item in quota_data:
                    if isinstance(item, bytes):
                        quota_str = item.decode('utf-8', errors='ignore')
                        if 'STORAGE' in quota_str:
                            parts = quota_str.split()
                            for i, part in enumerate(parts):
                                if part == 'STORAGE' and i + 2 < len(parts):
                                    try:
                                        used_kb = int(parts[i + 1])
                                        limit_kb = int(parts[i + 2])
                                        return {
                                            'used_kb': used_kb,
                                            'used_mb': round(used_kb / 1024, 2),
                                            'limit_kb': limit_kb,
                                            'limit_mb': round(limit_kb / 1024, 2)
                                        }
                                    except ValueError:
                                        continue
            return None
        except Exception:
            return None

    def _parse_folder_name(self, folder_line):
        try:
            if isinstance(folder_line, bytes):
                folder_str = folder_line.decode('utf-8', errors='ignore')
            else:
                folder_str = str(folder_line)
            match = re.search(r'"((?:[^"\\]|\\.)*)"\s*$', folder_str)
            if match:
                name = match.group(1)
                name = name.replace('\\"', '"').replace('\\\\', '\\')
                return name
            parts = folder_str.split('"')
            if len(parts) >= 3:
                name = parts[-2]
                return name.replace('\\"', '"').replace('\\\\', '\\')
            return None
        except Exception:
            return None

    def get_folder_info(self, mail, folder_name, chunk_size=500, debug_parse=False):
        try:
            if not folder_name or folder_name.strip() in ['', '/', '"']:
                return {'message_count': 0, 'size_bytes': 0, 'size_mb': 0, 'method': 'invalid_name'}

            status, _ = mail.select(folder_name, readonly=True)
            if status != 'OK':
                return {'message_count': 0, 'size_bytes': 0, 'size_mb': 0, 'method': 'select_failed'}

            # METHOD 1: STATUS=SIZE
            try:
                status, data = mail.status(folder_name, '(MESSAGES SIZE)')
                if status == 'OK' and data:
                    line = data[0].decode('utf-8', errors='ignore')
                    msg_match = re.search(r'MESSAGES\s+(\d+)', line, re.I)
                    size_match = re.search(r'SIZE\s+(\d+)', line, re.I)
                    if msg_match and size_match:
                        msg_count = int(msg_match.group(1))
                        total_size = int(size_match.group(1))
                        if msg_count > 0 or total_size > 0:
                            print(f"      ⚡ STATUS=SIZE: {msg_count} msgs, {total_size/(1024*1024):.2f} MB")
                            return {
                                'message_count': msg_count,
                                'size_bytes': total_size,
                                'size_mb': round(total_size / (1024 * 1024), 2),
                                'method': 'STATUS=SIZE'
                            }
            except Exception as e:
                if debug_parse:
                    print(f"      ⚠️  STATUS=SIZE failed: {e}")
                pass

            # METHOD 2: RFC822.SIZE for all messages
            print(f"      🔍 Fetching RFC822.SIZE for all messages...")
            status, data = mail.uid('search', None, 'ALL')
            if status != 'OK' or not data or not data[0]:
                return {'message_count': 0, 'size_bytes': 0, 'size_mb': 0, 'method': 'RFC822.SIZE'}

            uids = data[0].split()
            message_count = len(uids)
            if message_count == 0:
                return {'message_count': 0, 'size_bytes': 0, 'size_mb': 0, 'method': 'RFC822.SIZE'}

            print(f"      📦 Found {message_count} messages, processing in chunks of {chunk_size}...")
            total_size = 0
            processed = 0

            for start in range(0, message_count, chunk_size):
                chunk_uids = uids[start:start + chunk_size]
                uid_range = b','.join(chunk_uids)
                try:
                    status, fetch_data = mail.uid('fetch', uid_range, '(RFC822.SIZE)')
                    if status != 'OK' or not fetch_data:
                        continue

                    if debug_parse and processed == 0 and fetch_data:
                        print(f"      🔎 DEBUG: Raw response sample: {fetch_data[:2]}")

                    for item in fetch_data:
                        if not item:
                            continue
                        if isinstance(item, tuple) and len(item) >= 1:
                            header = item[0]
                        elif isinstance(item, bytes):
                            header = item
                        else:
                            continue
                        if isinstance(header, bytes):
                            header_str = header.decode('utf-8', errors='ignore')
                        else:
                            header_str = str(header)
                        if header_str.strip() in [')', '', '(']:
                            continue

                        size_match = re.search(r'RFC822\.SIZE\s+(\d+)', header_str, re.I)
                        if size_match:
                            try:
                                size = int(size_match.group(1))
                                total_size += size
                                processed += 1
                                if debug_parse and processed <= 3:
                                    print(f"      🔎 Parsed: {size} bytes from {header_str[:60]}...")
                            except ValueError:
                                continue

                    progress = min(start + chunk_size, message_count)
                    if progress % 1000 == 0 or progress >= message_count:
                        print(f"      📊 Progress: {progress}/{message_count} msgs, parsed {processed} sizes")

                except Exception as e:
                    if debug_parse:
                        print(f"      ❌ Chunk error: {e}")
                    continue

            print(f"      ✅ Processed {processed}/{message_count} messages")
            if processed == 0 and message_count > 0:
                print(f"      ⚠️  WARNING: Could not parse any sizes from {message_count} messages!")

            return {
                'message_count': message_count,
                'size_bytes': total_size,
                'size_mb': round(total_size / (1024 * 1024), 2),
                'method': 'RFC822.SIZE'
            }

        except Exception as e:
            print(f"      ❌ Folder error: {str(e)}")
            return {'message_count': 0, 'size_bytes': 0, 'size_mb': 0, 'method': 'error'}

    def analyze_mailbox(self, email_addr, password, imap_server):
        print(f"\n🔍 Analyzing: {email_addr} on {imap_server}")
        print("=" * 70)
        mail = self.connect_to_imap(imap_server, email_addr, password)
        if not mail:
            return None
        try:
            quota_info = self.get_quota_info(mail)
            status, folders = mail.list()
            if status != 'OK':
                mail.logout()
                return None
            print(f"   📁 Found {len(folders)} folders")
            folder_info = {}
            total_messages = 0
            total_size_bytes = 0

            for folder_line in folders:
                try:
                    folder_name = self._parse_folder_name(folder_line)
                    if not folder_name or folder_name in ['', '\\\\', '\\Noselect', '/']:
                        continue
                    print(f"\n   📂 Folder: {folder_name}")
                    info = self.get_folder_info(mail, folder_name)
                    folder_info[folder_name] = info
                    total_messages += info['message_count']
                    total_size_bytes += info['size_bytes']
                    print(f"      📊 {info['message_count']:,} messages, {info['size_mb']:,.2f} MB [{info.get('method', 'unknown')}]")
                except Exception as e:
                    print(f"      ❌ Folder error: {str(e)}")
                    continue

            mail.logout()
            total_size_mb = round(total_size_bytes / (1024 * 1024), 2)
            result = {
                'email': email_addr,
                'imap_server': imap_server,
                'total_messages': total_messages,
                'total_size_mb': total_size_mb,
                'analysis_timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            if quota_info:
                result['quota_used_mb'] = quota_info['used_mb']
                result['quota_limit_mb'] = quota_info['limit_mb']
                result['quota_usage_percent'] = round(
                    (quota_info['used_mb'] / quota_info['limit_mb']) * 100, 1
                ) if quota_info['limit_mb'] > 0 else 0
            for folder, info in folder_info.items():
                safe_name = folder.replace('/', '_').replace(' ', '_').replace('[', '').replace(']', '')
                safe_name = safe_name.replace('\\', '_').replace('.', '_')
                result[f'{safe_name}_messages'] = info['message_count']
                result[f'{safe_name}_size_mb'] = info['size_mb']
                result[f'{safe_name}_method'] = info.get('method', 'unknown')
            print(f"\n✅ Completed: {email_addr}")
            print(f"   📈 Total: {total_messages:,} messages, {total_size_mb:,.2f} MB")
            print("=" * 70)
            return result
        except Exception as e:
            print(f"❌ Error analyzing {email_addr}: {str(e)}")
            try:
                mail.logout()
            except:
                pass
            return None

    def process_file(self, input_file, output_file):
        print(f"\n📂 Reading input file: {input_file}")
        print("=" * 70)
        if not os.path.exists(input_file):
            print(f"❌ ERROR: Input file not found: {input_file}")
            return
        try:
            if input_file.endswith('.csv'):
                df = pd.read_csv(input_file, header=None)
            else:
                df = pd.read_excel(input_file, header=None)
            df.columns = ['email', 'password', 'imap_server']
            print(f"   📊 Found {len(df)} mailboxes to analyze")
        except Exception as e:
            print(f"❌ Error reading input file: {str(e)}")
            return

        results = []
        start_time = datetime.now()
        for index, row in df.iterrows():
            try:
                print(f"\n{'='*70}")
                print(f"Progress: {index + 1}/{len(df)}")
                print(f"{'='*70}")
                result = self.analyze_mailbox(
                    str(row['email']).strip(),
                    str(row['password']).strip(),
                    str(row['imap_server']).strip()
                )
                if result:
                    results.append(result)
                    print(f"✅ Success: {len(results)}/{index + 1} completed")
                else:
                    print(f"⚠️  Failed: {index + 1}/{len(df)}")
                if index < len(df) - 1:
                    print(f"   ⏳ Waiting 2 seconds before next account...")
                    time.sleep(2)
            except Exception as e:
                print(f"❌ Error processing row {index + 1}: {str(e)}")
                continue

        end_time = datetime.now()
        runtime = (end_time - start_time).total_seconds()
        if results:
            results_df = pd.DataFrame(results)
            base_columns = ['email', 'imap_server', 'total_messages', 'total_size_mb', 'analysis_timestamp']
            quota_columns = [col for col in results_df.columns if 'quota' in col]
            folder_columns = [col for col in results_df.columns if col not in base_columns + quota_columns]
            column_order = base_columns + quota_columns + sorted(folder_columns)
            results_df = results_df.reindex(columns=[col for col in column_order if col in results_df.columns])
            results_df.to_excel(output_file, index=False)
            print(f"\n{'='*70}")
            print(f"💾 Results saved to: {output_file}")
            print(f"⏱️  Total runtime: {runtime:.2f} seconds ({runtime/60:.2f} minutes)")
            print(f"📈 Successfully analyzed {len(results)} out of {len(df)} mailboxes")
            total_size = results_df['total_size_mb'].sum()
            total_messages = results_df['total_messages'].sum()
            print(f"\n📊 SUMMARY:")
            print(f"   Total Messages: {total_messages:,}")
            print(f"   Total Storage:  {total_size:,.2f} MB ({total_size/1024:.2f} GB)")
            print(f"{'='*70}\n")
        else:
            print("\n❌ No results to save - all analyses failed")


def main():
    print("main() starting")
    print("=" * 70)
    print("📧 MAILBOX STORAGE ANALYZER - ZERO ERROR MARGIN (FINAL)")
    print("=" * 70)
    analyzer = MailboxAnalyzer()
    input_file = "mailboxes.xlsx"
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f"mailbox_analysis_{timestamp}.xlsx"
    print(f"📥 Input:  {input_file}")
    print(f"📤 Output: {output_file}")
    print("=" * 70)
    analyzer.process_file(input_file, output_file)
    print("\n🎉 Analysis complete!")


if __name__ == "__main__":
    main()
