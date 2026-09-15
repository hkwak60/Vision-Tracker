import unittest
from dataclasses import replace
from vision_tracker import (
    DeepLearningTrainedInput, DeepLearningApplicationInput,
    validate_dl_trained_model, validate_dl_application,
)


class ChangeTypeTests(unittest.TestCase):
    def test_training_and_application_types_are_separate(self):
        training = DeepLearningTrainedInput(
            '2026-09-15 09:00', 'SEPA', 'v2', 'Overkill Training',
            ('1-1 Welding(-)',), 'Hojun Kwak', '',
        )
        application = DeepLearningApplicationInput(
            '2026-09-15 09:00', 'SEPA', 'v2', 'Model Update',
            ('1-1 Welding(-)',), 'Hojun Kwak', '',
        )
        for value in ['Overkill Training', 'Leakage Training', 'Overkill&Leakage', 'Dataset Major Change']:
            with self.subTest(value=value):
                self.assertEqual(validate_dl_trained_model(replace(training, change_type=value)), [])
                self.assertTrue(validate_dl_application(replace(application, change_type=value)))
        for value in ['Model Update', 'Model Revert']:
            with self.subTest(value=value):
                self.assertEqual(validate_dl_application(replace(application, change_type=value)), [])
                self.assertTrue(validate_dl_trained_model(replace(training, change_type=value)))
        for value in ['', 'Unknown']:
            self.assertTrue(validate_dl_trained_model(replace(training, change_type=value)))
            self.assertTrue(validate_dl_application(replace(application, change_type=value)))


if __name__ == '__main__':
    unittest.main()
