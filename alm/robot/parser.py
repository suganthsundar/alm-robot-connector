"""
Robot Output XML Parser for ALM update
"""
from xml.etree import ElementTree as ET


class RobotXMLParser(object):
    """ Robot output.xml file parser """

    ACTUAL_LOG_LEVEL = ('INFO', 'FAIL')

    def __init__(self, path):

        self.steps = []
        self.step_no = 0
        tree = ET.parse(path)
        for test in tree.iter(tag='test'):
            self.add_steps(test.attrib['name'], test)

    def add_steps(self, name, test):

        for kw in test.findall('kw'):
            if kw.findall('kw'):
                self.add_steps(name, kw)
            else:
                self.step_no += 1
                step = {}
                # step no, name & description
                step.update({"no": self.step_no})
                step.update({"name": name})
                step.update({"desc": kw.attrib['name']})

                # status
                status = kw.find('status').attrib['status']
                step.update({"status": status.capitalize() + 'ed'})

                # actual msg
                msgs = kw.findall('msg')
                msgList = []
                for msg in msgs:
                    if msg.attrib['level'] in self.ACTUAL_LOG_LEVEL:
                        msgList.append(msg.text)
                step.update({"actual": "\n".join(msgList)})

                self.steps.append(step)
