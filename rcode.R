# Load necessary libraries
library(readxl)
library(dplyr)
library(ggplot2)
library(forecast)
library(tidyr)

# Read the data from the Excel file
Panels_Per_Lot_2023 <- read_excel("Panels Per Lot 2023.xlsx")
View(Panels_Per_Lot_2023)

file_path <- "Panels Per Lot 2023.xlsx"
data <- read_excel(file_path, sheet = "Sheet1")

# Filter data for Lot 4
lot4_data <- data %>% filter(Lot == 4)

# Tidy the data
tidy_data <- lot4_data %>%
  select(Array, Modules, Tilt, Azimuth, Inverters, `Panels/Inverter`, `Strings/Inverter`, `Panels/String`, `Panel Power`, `Array Power/Inverter`, `Overload Ratio`) %>%
  mutate(
    Total_Power = `Panels/Inverter` * `Panel Power` * `Strings/Inverter`,
    Array = as.factor(Array)
  )

# Analyze the power generation
summary_stats <- tidy_data %>%
  group_by(Array) %>%
  summarise(
    Total_Modules = sum(Modules),
    Total_Power = sum(Total_Power),
    Avg_Tilt = mean(Tilt),
    Avg_Azimuth = mean(Azimuth)
  )

print(summary_stats)

# Plot the power generation
ggplot(tidy_data, aes(x = Array, y = Total_Power, fill = Array)) +
  geom_bar(stat = "identity") +
  labs(title = "Total Power Generation by Array", x = "Array", y = "Total Power (W)") +
  theme_minimal()

# Forecast power generation
# Assuming we have historical power generation data for forecasting
# For demonstration, we'll create a synthetic time series data
set.seed(123)
time_series_data <- ts(rnorm(365, mean = 5000, sd = 1000), frequency = 365, start = c(2023, 1))

# Plot the historical data
autoplot(time_series_data) +
  labs(title = "Historical Power Generation", x = "Time", y = "Power (W)")

# Fit an ARIMA model for forecasting
fit <- auto.arima(time_series_data)

# Forecast the next 30 days
forecast_data <- forecast(fit, h = 30)

# Plot the forecast
autoplot(forecast_data) +
  labs(title = "Power Generation Forecast", x = "Time", y = "Power (W)")

# Save the tidy data to a CSV file
write.csv(tidy_data, "tidy_power_generation_data.csv", row.names = FALSE)

