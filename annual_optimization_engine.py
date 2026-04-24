import random
import pandas as pd
import numpy as np

# Mocking the parts for the simulation (based on parts_master.xlsx)
PARTS = [
    {"id": f"CASE_{i:02d}", "w": w, "d": d, "h": h, "weight": random.randint(500, 2000)}
    for i, (w, d, h) in enumerate([
        (1490, 2260, 550), (1490, 2260, 730), (1490, 2260, 900), (1490, 2260, 1100),
        (1490, 2260, 1460), (1490, 2260, 1650), (1490, 1130, 550), (1490, 1130, 730),
        (1990, 2260, 550), (1990, 2260, 730), (1990, 2260, 900)
    ], 1)
]

CONTAINER_W, CONTAINER_D, CONTAINER_H = 12000, 2300, 2400
CONTAINER_MAX_WEIGHT = 15000
CONTAINER_VOLUME = CONTAINER_W * CONTAINER_D * CONTAINER_H

def generate_weekly_cargo():
    """Generates random cargo for one week (approx 5-15 containers worth)"""
    num_items = random.randint(100, 300)
    items = []
    for _ in range(num_items):
        p = random.choice(PARTS)
        items.append(p.copy())
    return items

def simulate_vanning(items, target_utilization=1.0):
    """
    Simulates packing items into containers.
    target_utilization: 0.55 to 0.85 for 'inefficient', 1.0 for 'optimized'
    """
    containers = []
    current_container_items = []
    current_weight = 0
    current_volume = 0
    
    # Sort items by volume (descending) for better packing efficiency
    sorted_items = sorted(items, key=lambda x: x['w']*x['d']*x['h'], reverse=True)
    
    for item in sorted_items:
        vol = item['w'] * item['d'] * item['h']
        weight = item['weight']
        
        # Check if item fits in current container based on target utilization
        # (Simplified: using volume and weight limits)
        if (current_volume + vol <= CONTAINER_VOLUME * target_utilization and 
            current_weight + weight <= CONTAINER_MAX_WEIGHT * target_utilization):
            current_container_items.append(item)
            current_volume += vol
            current_weight += weight
        else:
            # Start new container
            if current_container_items:
                containers.append({
                    "items": current_container_items,
                    "utilization_vol": current_volume / CONTAINER_VOLUME,
                    "utilization_weight": current_weight / CONTAINER_MAX_WEIGHT
                })
            current_container_items = [item]
            current_volume = vol
            current_weight = weight
            
    if current_container_items:
        containers.append({
            "items": current_container_items,
            "utilization_vol": current_volume / CONTAINER_VOLUME,
            "utilization_weight": current_weight / CONTAINER_MAX_WEIGHT
        })
        
    return containers

def run_annual_simulation():
    total_containers_inefficient = 0
    total_containers_optimized = 0
    
    results = []
    
    for week in range(1, 53):
        weekly_items = generate_weekly_cargo()
        
        # Inefficient: 55-85% utilization
        util = random.uniform(0.55, 0.85)
        containers_inf = simulate_vanning(weekly_items, target_utilization=util)
        
        # Optimized: 95-100% utilization (allowing some slack for real packing geometry)
        containers_opt = simulate_vanning(weekly_items, target_utilization=0.95)
        
        total_containers_inefficient += len(containers_inf)
        total_containers_optimized += len(containers_opt)
        
        results.append({
            "Week": week,
            "Items": len(weekly_items),
            "Inefficient_Count": len(containers_inf),
            "Optimized_Count": len(containers_opt),
            "Saved": len(containers_inf) - len(containers_opt)
        })
        
    print(f"--- Annual Simulation Results ---")
    print(f"Total Weeks: 52")
    print(f"Total Containers (Before): {total_containers_inefficient}")
    print(f"Total Containers (After): {total_containers_optimized}")
    print(f"Containers Saved: {total_containers_inefficient - total_containers_optimized}")
    reduction = (total_containers_inefficient - total_containers_optimized) / total_containers_inefficient * 100
    print(f"Reduction Rate: {reduction:.2f}%")
    
    return results

if __name__ == "__main__":
    run_annual_simulation()
