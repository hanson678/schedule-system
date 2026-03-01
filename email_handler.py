# -*- coding: utf-8 -*-
"""邮件处理 - 支持IMAP和Foxmail目录监控"""
import os
import re
import email
import imaplib
import tempfile
from datetime import datetime, timedelta
from email.header import decode_header


class EmailHandler:
    def __init__(self, config):
        self.config = config
        self.imap_server = config.get('imap_server', '')
        self.imap_port = config.get('imap_port', 993)
        self.email_account = config.get('email_account', '')
        self.email_password = config.get('email_password', '')
        self.foxmail_path = config.get('foxmail_path', '')
        self.upload_dir = os.path.join(os.path.dirname(__file__), 'uploads')
        os.makedirs(self.upload_dir, exist_ok=True)

    def fetch_recent_po_emails(self, days=3):
        """通过IMAP获取最近的PO相关邮件"""
        if not self.imap_server or not self.email_password:
            return {'error': '请先配置IMAP服务器和密码', 'emails': []}

        try:
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.email_account, self.email_password)
            mail.select('INBOX')

            # 搜索最近N天的邮件
            since = (datetime.now() - timedelta(days=days)).strftime('%d-%b-%Y')
            _, msg_nums = mail.search(None, f'(SINCE {since})')

            emails = []
            for num in msg_nums[0].split():
                _, msg_data = mail.fetch(num, '(RFC822)')
                raw = msg_data[0][1]
                msg = email.message_from_bytes(raw)

                subject = self._decode_header(msg['Subject'] or '')
                from_addr = self._decode_header(msg['From'] or '')
                date_str = msg['Date'] or ''

                # 检查是否PO相关
                is_po = bool(re.search(r'PO|4500\d{6}|Purchase Order|订单', subject, re.IGNORECASE))

                # 提取附件
                attachments = []
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        filename = part.get_filename()
                        if filename:
                            filename = self._decode_header(filename)
                            if filename.lower().endswith('.pdf'):
                                # 保存PDF
                                save_path = os.path.join(self.upload_dir, filename)
                                with open(save_path, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                attachments.append({
                                    'filename': filename,
                                    'path': save_path,
                                    'size': os.path.getsize(save_path)
                                })

                # 提取正文中的PO号
                body = self._get_body(msg)
                po_numbers = re.findall(r'4500\d{6}', body + subject)

                if is_po or po_numbers or attachments:
                    emails.append({
                        'subject': subject,
                        'from': from_addr,
                        'date': date_str,
                        'po_numbers': list(set(po_numbers)),
                        'attachments': attachments,
                        'body_preview': body[:500],
                        'is_new': self._detect_order_type(subject, body) == 'new',
                        'is_modify': self._detect_order_type(subject, body) == 'modify',
                        'is_cancel': self._detect_order_type(subject, body) == 'cancel',
                        'order_type': self._detect_order_type(subject, body)
                    })

            mail.logout()
            return {'emails': emails}

        except Exception as e:
            return {'error': str(e), 'emails': []}

    def _decode_header(self, header):
        """解码邮件头"""
        if not header:
            return ''
        decoded = decode_header(header)
        parts = []
        for content, charset in decoded:
            if isinstance(content, bytes):
                parts.append(content.decode(charset or 'utf-8', errors='replace'))
            else:
                parts.append(content)
        return ''.join(parts)

    def _get_body(self, msg):
        """获取邮件正文"""
        if msg.is_multipart():
            for part in msg.walk():
                ct = part.get_content_type()
                if ct == 'text/plain':
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or 'utf-8'
                    return payload.decode(charset, errors='replace')
                elif ct == 'text/html':
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or 'utf-8'
                    text = payload.decode(charset, errors='replace')
                    # 简单去HTML标签
                    text = re.sub(r'<[^>]+>', ' ', text)
                    return text
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or 'utf-8'
                return payload.decode(charset, errors='replace')
        return ''

    def _detect_order_type(self, subject, body):
        """检测订单类型: new/modify/cancel"""
        combined = (subject + ' ' + body).lower()

        cancel_kw = ['cancel', '取消', 'cancellation', '撤单']
        modify_kw = ['修改', 'change', 'modify', 'amendment', 'revision', 'update', '变更', '更改']
        new_kw = ['new po', '新单', 'new order', 'purchase order']

        for kw in cancel_kw:
            if kw in combined:
                return 'cancel'
        for kw in modify_kw:
            if kw in combined:
                return 'modify'
        for kw in new_kw:
            if kw in combined:
                return 'new'

        # 默认: 有PO号的视为新单
        if re.search(r'4500\d{6}', combined):
            return 'new'
        return 'unknown'

    def scan_foxmail_attachments(self):
        """扫描Foxmail附件目录（备用方案）"""
        attach_dir = os.path.join(self.foxmail_path, 'CommintAttachs')
        if not os.path.isdir(attach_dir):
            return []

        pdfs = []
        for root, dirs, files in os.walk(attach_dir):
            for f in files:
                if f.lower().endswith('.pdf') and ('po' in f.lower() or '4500' in f):
                    fp = os.path.join(root, f)
                    pdfs.append({
                        'filename': f,
                        'path': fp,
                        'size': os.path.getsize(fp),
                        'modified': datetime.fromtimestamp(os.path.getmtime(fp)).strftime('%Y-%m-%d %H:%M')
                    })
        return sorted(pdfs, key=lambda x: x['modified'], reverse=True)
