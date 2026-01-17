"""Canvas API helper functions for syncing assignments."""
import requests
from typing import List, Dict, Optional
from datetime import datetime


def parse_date_to_iso(date_str: str) -> str:
    """Parse various date formats to ISO 8601 YYYY-MM-DD format.
    
    Args:
        date_str: Date in format like "4/9/26", "2026-04-09", etc.
        
    Returns:
        ISO 8601 date string (YYYY-MM-DD)
    """
    if not date_str:
        return None
    
    # Try common formats
    formats = [
        '%Y-%m-%d',      # 2026-04-09
        '%m/%d/%y',      # 4/9/26
        '%m/%d/%Y',      # 4/9/2026
        '%Y/%m/%d',      # 2026/04/09
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            # If year is < 100, assume it's 20xx
            if dt.year < 100:
                dt = dt.replace(year=dt.year + 2000)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            continue
    
    # If all else fails, return original
    return date_str


class CanvasAPI:
    def __init__(self, api_key: str, canvas_url: str):
        """Initialize Canvas API client.
        
        Args:
            api_key: Canvas API access token
            canvas_url: Base URL for Canvas instance (e.g., https://canvas.instructure.com)
        """
        self.api_key = api_key
        self.canvas_url = canvas_url.rstrip('/')
        self.headers = {'Authorization': f'Bearer {api_key}'}
    
    def get_course_assignments(self, course_id: str) -> List[Dict]:
        """Fetch all assignments for a course.
        
        Args:
            course_id: Canvas course ID
            
        Returns:
            List of assignment dictionaries with id, name, due_at, etc.
        """
        url = f'{self.canvas_url}/api/v1/courses/{course_id}/assignments'
        params = {'per_page': 100}
        assignments = []
        
        while url:
            response = requests.get(url, headers=self.headers, params=params)
            response.raise_for_status()
            assignments.extend(response.json())
            
            # Check for pagination
            if 'next' in response.links:
                url = response.links['next']['url']
                params = {}  # params already in next URL
            else:
                url = None
        
        return assignments
    
    def update_assignment(self, course_id: str, assignment_id: str, 
                         name: Optional[str] = None, 
                         due_date: Optional[str] = None,
                         description: Optional[str] = None,
                         points: Optional[float] = None) -> Dict:
        """Update a Canvas assignment.
        
        Args:
            course_id: Canvas course ID
            assignment_id: Canvas assignment ID
            name: New assignment name (optional)
            due_date: New due date in ISO 8601 format (optional)
            description: New description (optional)
            points: New points possible (optional)
            
        Returns:
            Updated assignment dictionary
        """
        url = f'{self.canvas_url}/api/v1/courses/{course_id}/assignments/{assignment_id}'
        data = {'assignment': {}}
        
        if name is not None:
            data['assignment']['name'] = name
        if due_date is not None:
            # Parse and convert to ISO 8601 with timezone
            iso_date = parse_date_to_iso(due_date)
            if iso_date:
                # Add time component for Canvas
                data['assignment']['due_at'] = f'{iso_date}T23:59:00Z'
        if description is not None:
            data['assignment']['description'] = description
        if points is not None:
            data['assignment']['points_possible'] = points
        
        response = requests.put(url, headers=self.headers, json=data)
        response.raise_for_status()
        return response.json()
    
    def create_assignment(self, course_id: str, name: str, due_date: str,
                         description: Optional[str] = None,
                         points: Optional[float] = None) -> Dict:
        """Create a new Canvas assignment.
        
        Args:
            course_id: Canvas course ID
            name: Assignment name
            due_date: Due date in ISO 8601 format
            description: Assignment description (optional)
            points: Points possible (optional)
            
        Returns:
            Created assignment dictionary
        """
        url = f'{self.canvas_url}/api/v1/courses/{course_id}/assignments'
        data = {
            'assignment': {
                'name': name,
                'due_at': due_date,
                'published': False  # Create as unpublished by default
            }
        }
        
        if description is not None:
            data['assignment']['description'] = description
        if points is not None:
            data['assignment']['points_possible'] = points
        
        response = requests.post(url, headers=self.headers, json=data)
        response.raise_for_status()
        return response.json()
