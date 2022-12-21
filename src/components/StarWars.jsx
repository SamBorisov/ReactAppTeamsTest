import React from 'react'


// STAR WARS HERO GENERATOR---------------------------------------------------------

export default function StarWars() {

    const [starWarsData, setStarWarsData] = React.useState({})
    const [count, setCount] = React.useState(0)
    
    React.useEffect(function() {
        console.log("Effect ran")
        fetch("https://swapi.dev/api/people/" + count)
            .then(res => res.json())
            .then(data => setStarWarsData(data))
    }, [count])
    return (
        <div>
            <pre>{JSON.stringify(starWarsData, null, 2)}</pre>
            <button onClick={() => setCount(prevCount => prevCount + 1)}>Add</button>
            <h2>Click add to see next hero {count}</h2>
            
        </div>)
}
