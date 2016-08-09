"""
Robot Test Runner & Reports Update in ALM
"""
import os
import sys
import robot
import shutil
import time
import win32com

from alm.robot.parser import RobotXMLParser


class TestRunner(object):
    """ Robot Test Runner """

    def __init__(self, name, **alm):
        self.alm = alm
        self.options = {}
        self.console_log = None
        self.app_name = name

    def _get_outputdir(self, test_path):

        test_path_list = test_path.split('\\')
        test_name = test_path_list[len(test_path_list) - 1]
        test_name = test_name.split('.')[0]

        dir_config = {
            'root': os.path.expanduser('~'),
            'base': r'tests\reports\{0}'.format(self.app_name),
            'year': time.strftime('%Y'),
            'month': time.strftime('%m'),
            'date': time.strftime('%d'),
            'run': 'RUN_' + time.strftime('%H%M')
        }

        outputdir_format = '{root}\{base}\{year}\{month}\{date}\{run}'
        outputdir = outputdir_format.format(**dir_config)
        return os.path.abspath(outputdir)

    def run_test(self, test_path, **options):

        self.options = options
        status = None

        # sets temp dir if outputdir not provided
        if 'outputdir' not in self.options:
            self.options['outputdir'] = self._get_outputdir(test_path)

        # removes output dir if exists
        if os.path.exists(self.options['outputdir']):
            shutil.rmtree(self.options['outputdir'])
        os.makedirs(self.options['outputdir'])

        # logs console output to file on ALM execution
        if self.alm:
            log_path = os.path.join(self.options['outputdir'], 'console.log')
            self.console_log = open(log_path, 'a')
            sys.__stdout__ = sys.__stderr__ = sys.stdout = self.console_log

        # runs robot tests
        try:
            status = robot.run(test_path, **self.options)
            if self.alm:
                self.alm['TDOutput'].Print(
                    'Test Execution Path: %s' % (test_path))
                self.alm['TDOutput'].Print(
                    'Robot Run Completed with %s status' % (status))
        except Exception, err:
            if self.alm:
                self.alm['TDOutput'].Print(
                    'Error while Robot Run: %s' % (str(err)))
            else:
                print str(err)

        # closes console log file
        finally:
            if self.console_log:
                self.console_log.close()

    def add_steps(self):
        "Post Steps in ALM"

        output_file = os.path.join(self.options['outputdir'], 'output.xml')
        try:
            if not os.path.exists(output_file):
                return None
            output = RobotXMLParser(output_file)
            for step in output.steps:
                if self.alm:
                    self.alm['TDOutput'].Print(
                        "%s: %s" % (step['desc'], step['status']))
                    if not self.alm['Debug']:
                        self.alm['TDHelper'].AddStepToRun(
                            Name=step['name'],
                            Desc=step['desc'],
                            Expected='',
                            Actual=step['actual'],
                            Status=step['status'])
        except Exception, err:
            if self.alm:
                self.alm['TDOutput'].Print(
                    'Error while Robot Output Parsing: %s' % (str(err)))
            else:
                print str(err)

    def upload_attachments(self):
        """ Uploads Attachments in ALM """

        if self.alm:
            try:
                for rfile in os.listdir(self.options['outputdir']):
                    path = os.path.join(self.options['outputdir'], rfile)
                    self.alm['TDOutput'].Print("Uploading %s" % (path))
                    if not self.alm['Debug']:
                        self.alm['TDHelper'].UploadAttachment(
                            path, self.alm['CurrentRun'])
            except Exception, err:
                self.alm['TDOutput'].Print(
                    'Error while Uploading Reports: %s' % (str(err)))
