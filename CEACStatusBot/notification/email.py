from smtplib import SMTP_SSL
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

from .handle import NotificationHandle

class EmailNotificationHandle(NotificationHandle):
    def __init__(self,fromEmail:str,toEmail:str,emailPassword:str,hostAddress:str='') -> None:
        super().__init__()
        self.__fromEmail = fromEmail
        self.__toEmail = toEmail.split("|")
        self.__emailPassword = emailPassword
        self.__hostAddress = hostAddress or "smtp."+fromEmail.split("@")[1]
        if ':' in self.__hostAddress:
            [addr, port] = self.__hostAddress.split(':')
            self.__hostAddress = addr
            self.__hostPort = int(port)
        else:
            self.__hostPort = 0

    def _build_visual_content(self, data_dict):
        """Transform visa status dictionary into visual HTML representation"""
        
        # Determine color scheme based on current status
        status_value = data_dict.get('status', 'Unknown')
        if status_value == 'Issued':
            banner_tone = '#10b981'  # green shade
            icon_symbol = '✓'
        elif status_value == 'Refused':
            banner_tone = '#ef4444'  # red shade
            icon_symbol = '✗'
        elif 'Administrative Processing' in status_value:
            banner_tone = '#f59e0b'  # amber shade
            icon_symbol = '⏳'
        else:
            banner_tone = '#6366f1'  # indigo shade
            icon_symbol = '●'
        
        # Build structured visual layout
        visual_parts = []
        visual_parts.append('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">')
        visual_parts.append('<meta name="viewport" content="width=device-width, initial-scale=1.0">')
        visual_parts.append('<style>')
        visual_parts.append('body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; ')
        visual_parts.append('background-color: #f3f4f6; margin: 0; padding: 20px; }')
        visual_parts.append('.container { max-width: 600px; margin: 0 auto; background: white; ')
        visual_parts.append('border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }')
        visual_parts.append('.header-banner { background: linear-gradient(135deg, ' + banner_tone + ' 0%, ')
        visual_parts.append(banner_tone + 'dd 100%); color: white; padding: 32px 24px; text-align: center; }')
        visual_parts.append('.status-icon { font-size: 48px; margin-bottom: 12px; }')
        visual_parts.append('.status-title { font-size: 28px; font-weight: 700; margin: 0; }')
        visual_parts.append('.content-area { padding: 32px 24px; }')
        visual_parts.append('.info-row { display: flex; padding: 16px 0; border-bottom: 1px solid #e5e7eb; }')
        visual_parts.append('.info-label { font-weight: 600; color: #4b5563; min-width: 160px; }')
        visual_parts.append('.info-value { color: #1f2937; flex: 1; word-wrap: break-word; }')
        visual_parts.append('.description-box { margin-top: 24px; padding: 20px; background: #f9fafb; ')
        visual_parts.append('border-left: 4px solid ' + banner_tone + '; border-radius: 4px; }')
        visual_parts.append('.description-box p { margin: 0; color: #374151; line-height: 1.6; }')
        visual_parts.append('.footer-note { margin-top: 24px; padding-top: 24px; border-top: 2px solid #e5e7eb; ')
        visual_parts.append('text-align: center; color: #6b7280; font-size: 13px; }')
        visual_parts.append('</style></head><body><div class="container">')
        
        # Header section with status
        visual_parts.append('<div class="header-banner">')
        visual_parts.append(f'<div class="status-icon">{icon_symbol}</div>')
        visual_parts.append(f'<h1 class="status-title">{status_value}</h1>')
        visual_parts.append('</div>')
        
        # Content area with details
        visual_parts.append('<div class="content-area">')
        
        # Application number row
        app_num = data_dict.get('application_num_origin', 'N/A')
        visual_parts.append('<div class="info-row">')
        visual_parts.append('<span class="info-label">Application Number:</span>')
        visual_parts.append(f'<span class="info-value">{app_num}</span>')
        visual_parts.append('</div>')
        
        # Visa type row
        v_type = data_dict.get('visa_type', 'N/A')
        visual_parts.append('<div class="info-row">')
        visual_parts.append('<span class="info-label">Visa Category:</span>')
        visual_parts.append(f'<span class="info-value">{v_type}</span>')
        visual_parts.append('</div>')
        
        # Case created row
        created_date = data_dict.get('case_created', 'N/A')
        visual_parts.append('<div class="info-row">')
        visual_parts.append('<span class="info-label">Case Created:</span>')
        visual_parts.append(f'<span class="info-value">{created_date}</span>')
        visual_parts.append('</div>')
        
        # Last updated row
        updated_date = data_dict.get('case_last_updated', 'N/A')
        visual_parts.append('<div class="info-row">')
        visual_parts.append('<span class="info-label">Last Updated:</span>')
        visual_parts.append(f'<span class="info-value">{updated_date}</span>')
        visual_parts.append('</div>')
        
        # Query time row
        query_time = data_dict.get('time', 'N/A')
        visual_parts.append('<div class="info-row">')
        visual_parts.append('<span class="info-label">Query Time:</span>')
        visual_parts.append(f'<span class="info-value">{query_time}</span>')
        visual_parts.append('</div>')
        
        # Description box
        desc_text = data_dict.get('description', '')
        if desc_text:
            visual_parts.append('<div class="description-box">')
            visual_parts.append(f'<p>{desc_text}</p>')
            visual_parts.append('</div>')
        
        # Footer
        visual_parts.append('<div class="footer-note">')
        visual_parts.append('This notification was generated by CEACStatusBot')
        visual_parts.append('</div>')
        
        visual_parts.append('</div></div></body></html>')
        
        return ''.join(visual_parts)

    def send(self,result):
        
        # {'success': True, 'visa_type': 'NONIMMIGRANT VISA APPLICATION', 'status': 'Issued', 'case_created': '30-Aug-2022', 'case_last_updated': '19-Oct-2022', 'description': 'Your visa is in final processing. If you have not received it in more than 10 working days, please see the webpage for contact information of the embassy or consulate where you submitted your application.', 'application_num': '***'}

        mail_title = '[CEACStatusBot] {} : {}'.format(result["application_num_origin"],result['status'])
        mail_content = self._build_visual_content(result)

        msg = MIMEMultipart()
        msg["Subject"] = Header(mail_title,'utf-8')
        msg["From"] = self.__fromEmail
        msg['To'] = ";".join(self.__toEmail)
        msg.attach(MIMEText(mail_content,'html','utf-8'))

        smtp = SMTP_SSL(self.__hostAddress, self.__hostPort) # ssl登录
        print(smtp.login(self.__fromEmail,self.__emailPassword))
        print(smtp.sendmail(self.__fromEmail,self.__toEmail,msg.as_string()))
        smtp.quit()