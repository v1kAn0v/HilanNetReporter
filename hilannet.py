from main import Robot


class HilanNet(Robot):

    def payslip(self, attr):
        return super().wait_for_id(attr)
