import os
import tempfile
import unittest
from unittest.mock import patch


if 'DATABASE_PATH' not in os.environ:
    _temp_dir = tempfile.mkdtemp(prefix='outlookEmail-mailnest-')
    os.environ['DATABASE_PATH'] = os.path.join(_temp_dir, 'test.db')
if 'SECRET_KEY' not in os.environ:
    os.environ['SECRET_KEY'] = 'test-secret-key'

import web_outlook_app


class MailNestTestCase(unittest.TestCase):
    def setUp(self):
        self.app = web_outlook_app.app
        self.app.config['TESTING'] = True
        self.app.config['WTF_CSRF_ENABLED'] = False
        self.client = self.app.test_client()

        with self.client.session_transaction() as sess:
            sess['logged_in'] = True

        with self.app.app_context():
            web_outlook_app.init_db()
            db = web_outlook_app.get_db()
            db.execute('DELETE FROM temp_email_tags')
            db.execute('DELETE FROM tags')
            db.execute('DELETE FROM temp_email_messages')
            db.execute('DELETE FROM temp_emails')
            db.execute("UPDATE settings SET value = ? WHERE key = 'mailnest_api_key'", ('test-mailnest-key',))
            db.execute("UPDATE settings SET value = ? WHERE key = 'mailnest_base_url'", ('https://mailnest.top',))
            db.commit()

    def sample_order(self, email='demo@outlook.com', sale_mode='temporary', **overrides):
        order = {
            'id': 'order-1',
            'email': email,
            'sale_mode': sale_mode,
            'project_code': 'claude001',
            'project_name': 'Claude',
            'price': '0.018',
            'status': 'holding',
            'billing_status': 'pending',
            'started_at': '2026-06-10T12:00:00+08:00',
            'expired_at': '2026-06-10T12:20:00+08:00',
            'ended_at': None,
        }
        order.update(overrides)
        return order


class MailNestRequestTests(MailNestTestCase):
    def test_mailnest_request_success_and_business_error(self):
        with self.app.app_context():
            with patch.object(web_outlook_app.requests, 'get') as mock_get:
                mock_get.return_value.status_code = 200
                mock_get.return_value.json.return_value = {
                    'code': '00000',
                    'msg': '',
                    'data': {'balance': '1.000', 'frozen_balance': '0', 'available_balance': '1.000'},
                }
                result = web_outlook_app.mailnest_request('GET', '/api/v1/balance')
            self.assertTrue(result['success'])
            self.assertEqual(result['data']['available_balance'], '1.000')

            with patch.object(web_outlook_app.requests, 'post') as mock_post:
                mock_post.return_value.status_code = 200
                mock_post.return_value.json.return_value = {
                    'code': 'D0001',
                    'msg': '余额不足',
                    'data': None,
                }
                result = web_outlook_app.mailnest_request(
                    'POST',
                    '/api/v1/email/temporary/buy',
                    json_data={'project_code': 'claude001', 'count': 1},
                )
            self.assertFalse(result['success'])
            self.assertEqual(result['code'], 'D0001')
            self.assertEqual(result['error'], '余额不足')

    def test_mailnest_request_unauthorized(self):
        with self.app.app_context():
            with patch.object(web_outlook_app.requests, 'get') as mock_get:
                mock_get.return_value.status_code = 401
                mock_get.return_value.json.return_value = {'code': '401', 'msg': 'unauthorized'}
                result = web_outlook_app.mailnest_request('GET', '/api/v1/balance')
            self.assertFalse(result['success'])
            self.assertIn('API Key', result['error'])

    def test_mailnest_request_requires_api_key(self):
        with self.app.app_context():
            web_outlook_app.set_setting('mailnest_api_key', '')
            result = web_outlook_app.mailnest_request('GET', '/api/v1/balance')
            self.assertFalse(result['success'])
            self.assertIn('未配置迈巢 API Key', result['error'])


class MailNestGenerateTests(MailNestTestCase):
    def test_generate_temporary_requires_project_code(self):
        response = self.client.post('/api/temp-emails/generate', json={
            'provider': 'mailnest',
            'sale_mode': 'temporary',
            'count': 1,
        })
        payload = response.get_json()
        self.assertFalse(payload['success'])
        self.assertIn('项目', payload['error'])

    def test_generate_temporary_saves_emails(self):
        order = self.sample_order()
        with patch.object(web_outlook_app, 'mailnest_buy_emails', return_value={
            'success': True,
            'sale_mode': 'temporary',
            'orders': [order],
        }):
            response = self.client.post('/api/temp-emails/generate', json={
                'provider': 'mailnest',
                'sale_mode': 'temporary',
                'project_code': 'claude001',
                'count': 1,
            })
        payload = response.get_json()
        self.assertTrue(payload['success'], payload)
        self.assertEqual(payload['emails'], ['demo@outlook.com'])
        with self.app.app_context():
            saved = web_outlook_app.get_temp_email_by_address('demo@outlook.com')
            self.assertEqual(saved['provider'], 'mailnest')
            self.assertEqual(saved['mailnest_order_id'], 'order-1')
            self.assertEqual(saved['mailnest_project_code'], 'claude001')
            self.assertEqual(saved['mailnest_sale_mode'], 'temporary')

    def test_generate_exclusive_and_batch(self):
        orders = [
            self.sample_order('one@outlook.com', 'exclusive', id='ex-1', project_code=None, project_name=None),
            self.sample_order('two@outlook.com', 'exclusive', id='ex-2', project_code=None, project_name=None),
        ]
        with patch.object(web_outlook_app, 'mailnest_buy_emails', return_value={
            'success': True,
            'sale_mode': 'exclusive',
            'orders': orders,
        }) as buy_mock:
            response = self.client.post('/api/temp-emails/generate-batch', json={
                'provider': 'mailnest',
                'sale_mode': 'exclusive',
                'count': 2,
            })
        payload = response.get_json()
        self.assertTrue(payload['success'], payload)
        self.assertEqual(payload['created_count'], 2)
        self.assertEqual(buy_mock.call_args.kwargs.get('sale_mode') or buy_mock.call_args[0][0], 'exclusive')
        with self.app.app_context():
            one = web_outlook_app.get_temp_email_by_address('one@outlook.com')
            self.assertEqual(one['mailnest_sale_mode'], 'exclusive')

    def test_generate_rejects_invalid_count(self):
        response = self.client.post('/api/temp-emails/generate-batch', json={
            'provider': 'mailnest',
            'sale_mode': 'exclusive',
            'count': 0,
        })
        payload = response.get_json()
        self.assertFalse(payload['success'])
        self.assertIn('数量', payload['error'])


class MailNestMessageTests(MailNestTestCase):
    def test_receive_saves_messages_and_marks_charged(self):
        with self.app.app_context():
            self.assertTrue(web_outlook_app.add_temp_email(
                'demo@outlook.com',
                provider='mailnest',
                mailnest_order_id='order-1',
                mailnest_sale_mode='temporary',
                mailnest_billing_status='pending',
                mailnest_status='holding',
            ))

        with patch.object(web_outlook_app, 'mailnest_receive_messages', return_value={
            'success': True,
            'messages': [{
                'id': 'msg-1',
                'order_id': 'order-1',
                'email': 'demo@outlook.com',
                'subject': 'Your verification code',
                'from_email': 'no-reply@anthropic.com',
                'body_preview': 'Your code is 123456',
                'body': '<p>Your verification code is 123456.</p>',
                'body_type': 'html',
                'code_match': '123456',
                'received_at': '2026-06-10T12:03:00+08:00',
            }],
        }):
            response = self.client.get('/api/temp-emails/demo@outlook.com/messages')

        payload = response.get_json()
        self.assertTrue(payload['success'], payload)
        self.assertEqual(payload['method'], 'MailNest')
        self.assertEqual(payload['count'], 1)
        self.assertIn('123456', payload['emails'][0]['body_preview'])

        with self.app.app_context():
            saved = web_outlook_app.get_temp_email_by_address('demo@outlook.com')
            self.assertEqual(saved['mailnest_billing_status'], 'charged')
            msg = web_outlook_app.get_temp_email_message_by_id('msg-1')
            self.assertIsNotNone(msg)
            self.assertEqual(msg['subject'], 'Your verification code')

    def test_receive_pending_falls_back_to_cache(self):
        with self.app.app_context():
            self.assertTrue(web_outlook_app.add_temp_email(
                'demo@outlook.com',
                provider='mailnest',
                mailnest_order_id='order-1',
                mailnest_sale_mode='temporary',
            ))
            web_outlook_app.save_temp_email_messages('demo@outlook.com', [{
                'id': 'cached-1',
                'from_address': 'cached@example.com',
                'subject': 'cached',
                'content': 'old mail',
                'html_content': '',
                'has_html': False,
                'timestamp': 1,
            }])

        with patch.object(web_outlook_app, 'mailnest_receive_messages', return_value={
            'success': True,
            'messages': [],
            'pending': True,
            'error': '暂未取到匹配邮件，请稍后再试',
        }):
            response = self.client.get('/api/temp-emails/demo@outlook.com/messages')

        payload = response.get_json()
        self.assertTrue(payload['success'], payload)
        self.assertEqual(payload['count'], 1)
        self.assertEqual(payload['emails'][0]['id'], 'cached-1')
        self.assertTrue(payload['pending'])

    def test_delete_releases_pending_but_skips_charged(self):
        with self.app.app_context():
            self.assertTrue(web_outlook_app.add_temp_email(
                'pending@outlook.com',
                provider='mailnest',
                mailnest_billing_status='pending',
            ))
            self.assertTrue(web_outlook_app.add_temp_email(
                'charged@outlook.com',
                provider='mailnest',
                mailnest_billing_status='charged',
            ))

        with patch.object(web_outlook_app, 'mailnest_release_email', return_value={'success': True}) as release_mock:
            pending_response = self.client.delete('/api/temp-emails/pending@outlook.com')
            charged_response = self.client.delete('/api/temp-emails/charged@outlook.com')

        self.assertTrue(pending_response.get_json()['success'])
        self.assertTrue(charged_response.get_json()['success'])
        self.assertEqual(release_mock.call_count, 1)
        self.assertEqual(release_mock.call_args[0][0], 'pending@outlook.com')


class MailNestImportAndSettingsTests(MailNestTestCase):
    def test_products_and_balance_endpoints(self):
        with patch.object(web_outlook_app, 'mailnest_get_products', return_value={
            'success': True,
            'temporary': [{'code': 'claude001', 'name': 'Claude', 'price': '0.018', 'stock': 10}],
            'exclusive': {'stock': 5, 'price': '0.055'},
        }), patch.object(web_outlook_app, 'mailnest_get_balance', return_value={
            'success': True,
            'balance': '1.000',
            'frozen_balance': '0.055',
            'available_balance': '0.945',
        }):
            products = self.client.get('/api/mailnest/products').get_json()
            balance = self.client.get('/api/mailnest/balance').get_json()

        self.assertTrue(products['success'])
        self.assertEqual(products['temporary'][0]['code'], 'claude001')
        self.assertEqual(products['balance']['available_balance'], '0.945')
        self.assertTrue(balance['success'])
        self.assertEqual(balance['available_balance'], '0.945')

    def test_import_sync_and_manual_import(self):
        with patch.object(web_outlook_app, 'mailnest_list_orders', side_effect=[
            {'success': True, 'items': [self.sample_order('temp@outlook.com')], 'total': 1},
            {'success': True, 'items': [self.sample_order('ex@outlook.com', 'exclusive', id='ex-1')], 'total': 1},
        ]):
            response = self.client.post('/api/temp-emails/import-mailnest', json={})
        payload = response.get_json()
        self.assertTrue(payload['success'], payload)
        self.assertEqual(payload['added_count'], 2)

        manual = self.client.post('/api/temp-emails/import', json={
            'provider': 'mailnest',
            'account_string': 'manual@outlook.com',
        }).get_json()
        self.assertTrue(manual['success'], manual)
        with self.app.app_context():
            saved = web_outlook_app.get_temp_email_by_address('manual@outlook.com')
            self.assertEqual(saved['provider'], 'mailnest')

    def test_settings_round_trip(self):
        response = self.client.put('/api/settings', json={
            'mailnest_api_key': 'new-key',
            'mailnest_base_url': 'https://mailnest.top/',
        })
        self.assertTrue(response.get_json()['success'])
        settings = self.client.get('/api/settings').get_json()['settings']
        self.assertEqual(settings['mailnest_api_key'], 'new-key')
        self.assertEqual(settings['mailnest_base_url'], 'https://mailnest.top')

    def test_export_includes_mailnest_section(self):
        with self.app.app_context():
            self.assertTrue(web_outlook_app.add_temp_email(
                'demo@outlook.com',
                provider='mailnest',
                mailnest_sale_mode='temporary',
            ))
            temp_group_id = web_outlook_app.get_temp_email_group_id()
            export_result = web_outlook_app.build_group_export_content([temp_group_id])
        content = '\n'.join(export_result['lines'])
        self.assertIn('[mailnest]', content)
        self.assertIn('demo@outlook.com', content)


if __name__ == '__main__':
    unittest.main()
