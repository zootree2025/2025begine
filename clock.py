    def draw_hands(self, hours, minutes, seconds):
        # Draw hour hand
        angle = math.pi / 6 * (hours - 3 + minutes / 60.0)
        x = self.center_x + self.radius * 0.5 * math.cos(angle)
        y = self.center_y + self.radius * 0.5 * math.sin(angle)
        self.canvas.create_line(self.center_x, self.center_y, x, y, width=4, fill='black')

        # Draw minute hand
        angle = math.pi / 30 * (minutes - 15 + seconds / 60.0)
        x = self.center_x + self.radius * 0.75 * math.cos(angle)
        y = self.center_y + self.radius * 0.75 * math.sin(angle)
        self.canvas.create_line(self.center_x, self.center_y, x, y, width=2, fill='black')

        # Draw second hand
        angle = math.pi / 30 * (seconds - 15)
        x = self.center_x + self.radius * 0.85 * math.cos(angle)
        y = self.center_y + self.radius * 0.85 * math.sin(angle)
        self.canvas.create_line(self.center_x, self.center_y, x, y, width=1, fill='red')
