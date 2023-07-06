# nmo-bibliometric-analysis
Back-end API for Bibliometric Analysis of NMO Describing, Citing and Using Literature. </br></br>
#Build and run the container: </br> 
docker-compose up -d --build </br></br>
#Connect backend to network: </br> 
docker network connect nmo_network nmo_bibliometric_analysis_web_1 </br></br>
#connect db to network </br>
docker network connect nmo_network nmo_bibliometric_analysis_db_1 </br></br>
#inspect network to ensure front-end and back-end connected to network </br>
docker network inspect nmo_network
