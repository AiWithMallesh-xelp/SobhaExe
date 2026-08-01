import threading
import time
import unittest

from automation import AutomationController, AutomationStoppedByUser, _automation_checkpoint


class FakePage:
    def is_closed(self):
        return False

    def evaluate(self, *_args):
        return None

    def wait_for_timeout(self, milliseconds):
        time.sleep(milliseconds / 1000)


class AutomationControllerTests(unittest.TestCase):
    def test_pause_resume_and_quit(self):
        controller = AutomationController()
        self.assertFalse(controller.status()["paused"])
        self.assertTrue(controller.request("pause")["paused"])
        checkpoint_complete = threading.Event()
        worker = threading.Thread(
            target=lambda: (_automation_checkpoint(FakePage(), controller), checkpoint_complete.set())
        )
        worker.start()
        time.sleep(0.05)
        self.assertFalse(checkpoint_complete.is_set())

        self.assertFalse(controller.request("resume")["paused"])
        worker.join(1)
        self.assertTrue(checkpoint_complete.is_set())

        controller.complete()
        self.assertTrue(controller.status()["completed"])
        self.assertTrue(controller.request("quit")["quitting"])
        with self.assertRaises(AutomationStoppedByUser):
            controller.checkpoint()


if __name__ == "__main__":
    unittest.main()
