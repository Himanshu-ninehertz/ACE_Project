const COMPONENT_IMAGES = {
    flexDuctwork: '/static/images/components/flex_ductwork.svg',
    bend250mm: '/static/images/components/bend_250mm.svg',
    extractLouvre: '/static/images/components/extract_louvre.svg',
    extractGrilles: '/static/images/components/extract_grilles.svg',
    vcd: '/static/images/components/vcd.svg',
    extractFans: '/static/images/components/extract_fans.svg',
    attenuators: '/static/images/components/attenuators.svg',
    reducers: '/static/images/components/reducers.svg',
    shoes: '/static/images/components/shoes.svg',
    hru: '/static/images/components/hru.svg'
};

function getComponentImage(type) {
    return COMPONENT_IMAGES[type] || '';
}

// Example usage:
// const imgElement = document.createElement('img');
// imgElement.src = getComponentImage('flexDuctwork');
// imgElement.alt = 'Flex Ductwork';