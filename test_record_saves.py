import io
import threading
import time
import unittest
from unittest.mock import patch
import app
import vision_tracker as vt

class FakeUI:
    online_sync_running = False
    record_save_running = False
    def __init__(self):
        self.callbacks = []
        self.states = []
        self.main = threading.get_ident()
    def winfo_children(self): return []
    def after(self, delay, callback):
        assert threading.get_ident() == self.main
        self.callbacks.append(callback)
    def set_sync_status(self, *args): self.states.append(args)
    def refresh_version_history(self): pass
    def refresh_deep_learning_models(self): pass
    def drain(self):
        deadline = time.monotonic() + 3
        while self.record_save_running and time.monotonic() < deadline:
            if self.callbacks: self.callbacks.pop(0)()
            time.sleep(.005)
        assert not self.record_save_running

class SaveTests(unittest.TestCase):
    def test_slow_save_keeps_main_thread_free_and_blocks_duplicates(self):
        ui = FakeUI()
        gate = threading.Event()
        started = threading.Event()
        completed = []
        def work():
            assert threading.get_ident() != ui.main
            started.set()
            gate.wait(2)
        with patch.object(app, 'online_mode', return_value=True):
            app.VisionIssueApp.run_record_save(ui, work, lambda: completed.append(threading.get_ident()))
            self.assertTrue(started.wait(1))
            self.assertTrue(ui.record_save_running)
            app.VisionIssueApp.run_record_save(ui, lambda: self.fail('duplicate'), lambda: None)
            gate.set()
            ui.drain()
        self.assertEqual(completed, [ui.main])
    def test_failure_releases_save_and_does_not_report_success(self):
        ui = FakeUI()
        def fail(): raise ValueError('failed')
        with patch.object(app,'online_mode',return_value=True), patch.object(app.messagebox,'showerror') as error:
            app.VisionIssueApp.run_record_save(ui, fail, lambda: self.fail('false success'))
            ui.drain()
            error.assert_called_once()
    def test_save_waits_for_active_sync(self):
        ui = FakeUI()
        ui.online_sync_running = True
        work = threading.Event()
        with patch.object(app,'online_mode',return_value=True):
            app.VisionIssueApp.run_record_save(ui, work.set, lambda: None)
            self.assertFalse(work.is_set())
            ui.online_sync_running = False
            ui.drain()
            self.assertTrue(work.is_set())
    def test_http_error_does_not_display_html(self):
        error = vt._urllib_error.HTTPError('https://example.invalid',404,'Not found',{},io.BytesIO(b'<html>private page</html>'))
        with patch.object(vt,'_load_online_config',return_value={'apps_script_url':'https://example.invalid','api_token':'test'}), patch.object(vt._urllib_request,'urlopen',side_effect=error):
            with self.assertRaises(ValueError) as result: vt._online_request('ping')
        self.assertIn('404',str(result.exception))
        self.assertNotIn('<html>',str(result.exception))
        self.assertLess(len(str(result.exception)),300)
    def test_empty_cache_never_requests_network(self):
        with patch.object(vt,'_online_enabled',return_value=True), patch.object(vt,'_online_rows',return_value=[]), patch.object(vt,'_online_request') as request:
            self.assertEqual(vt.search_issues(),[])
            request.assert_not_called()

if __name__ == '__main__': unittest.main()
