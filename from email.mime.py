import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = ''
mail.Subject = 'test email'
mail.Body = 'Message body'
mail.HTMLBody = '<h2>this is a test</h2>' #this field is optional


mail.Send()