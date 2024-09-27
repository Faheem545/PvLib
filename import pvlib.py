import pvlib
import pandas as pd 
import matplotlib.pyplot as plt

# Set Location Class
location = pvlib.location.Location(latitude=37.978136, longitude=-121.312537, name='Lot 4', altitude=6, tz='Etc/GMT+7')

# Load data from the Excel file
data = pd.read_excel('Panels Per Lot 2023.xlsx', sheet_name='Sheet1')

# Extract relevant data for Lot 4
lot4_data = data[data['Lot'] == 4]

# Set Array Mounting Class based on the data
mount = [
    pvlib.pvsystem.FixedMount(surface_tilt=7, surface_azimuth=162),
    pvlib.pvsystem.FixedMount(surface_tilt=7, surface_azimuth=252),
    pvlib.pvsystem.FixedMount(surface_tilt=7, surface_azimuth=135),
    pvlib.pvsystem.FixedMount(surface_tilt=7, surface_azimuth=137),
    pvlib.pvsystem.FixedMount(surface_tilt=7, surface_azimuth=139),
    pvlib.pvsystem.FixedMount(surface_tilt=7, surface_azimuth=118)   
]

# Get the cec_modules and inverters databases
SAM_URL = "CEC Modules.csv"
cec_modules = pvlib.pvsystem.retrieve_sam(path=SAM_URL) 
cec_modules = cec_modules.drop('Manufacturer', axis=0)
cec_modules = cec_modules.loc[:, ~cec_modules.columns.duplicated()]
cec_inverters = pvlib.pvsystem.retrieve_sam('cecinverter')

# Get the module and inverter specifications
module_parameters = cec_modules['Hanwha_Q_CELLS_Q_PEAK_DUO_L_G6_3_425']
inverter_parameters = cec_inverters['Huawei_Technologies_Co___Ltd___SUN2000_40KTL_US__480V_']

# Set inverter parameters
num_inverters = 12
inverter_parameters.Vdco = 1000
inverter_parameters.Vdcmax = 1000
inverter_parameters.Idcmax = 88
inverter_parameters.Mppt_high = 1000
inverter_parameters.Mppt_low = 200
inverter_parameters.Paco = 36600.0

# Set the temperature model parameters
temperature_model_parameters = pvlib.temperature.TEMPERATURE_MODEL_PARAMETERS['sapm']['open_rack_glass_glass']

# Create arrays based on the data
arrays = []
for index, row in lot4_data.iterrows():
    mount_index = int(row['Array']) - 1
    arrays.append(
        pvlib.pvsystem.Array(
            mount=mount[mount_index],
            module_parameters=module_parameters,
            modules_per_string=int(row['Panels/String']),
            temperature_model_parameters=temperature_model_parameters,
            strings=int(row['Strings/Inverter'])
        )
    )

# Set the PV System class
losses_parameters = dict(shading=0, availability=0)

systems = [
    pvlib.pvsystem.PVSystem(arrays=[arrays], inverter_parameters=inverter_parameters, losses_parameters=losses_parameters),
    pvlib.pvsystem.PVSystem(arrays=[arrays], inverter_parameters=inverter_parameters, losses_parameters=losses_parameters)
]

# Create ModelChain objects
mc = [
    pvlib.modelchain.ModelChain(systems, location, aoi_model='physical', spectral_model='no_loss'),
    pvlib.modelchain.ModelChain(systems, location, aoi_model='physical', spectral_model='no_loss')
]

# Define times and get clearsky data
times = pd.date_range('2023-07-01', freq='15min', periods=4*24, tz=location.tz)
clearsky = location.get_clearsky(times)

# Run the model
mc.run_model(clearsky)
mc.run_model(clearsky)

# Plot results
mc.results.ac.plot(label='Inverter 1')
mc.results.ac.plot(label='Inverter 2')
plt.ylabel('System Output (kW)')
plt.legend()
plt.show()

# Calculate production
prod = 8 * mc.results.ac + 4 * mc.results.ac
prod.plot(label='Production')
plt.ylabel('System Output (kW)')
plt.legend()
plt.show()

# Save production data to CSV
prod.to_csv('pvlib_july1.csv')

# Fetch weather data (commented out for now)
# api_key = 'lNchMCSOEzggsbHeDy7kCAkg9VC0UrUJEVyx0uFZ'
# email = 'dmueller@pacific.edu'
# keys = ['ghi', 'dni', 'dhi', 'temp_air', 'wind_speed', 'albedo', 'precipitable_water']
# psm3, psm3_metadata = pvlib.iotools.get_psm3(
#     latitude=37.978136, 
#     longitude=-121.312537, 
#     api_key=api_key,
#     email=email, 
#     interval=15, 
#     names=2022,                                         
#     map_variables=True, 
#     leap_day=True,
#     attributes=keys
# )
