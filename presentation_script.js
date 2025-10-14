// Credora Capital Presentation Script

let currentSlide = 1;
const totalSlides = 15; // Update this as we add more slides

// Initialize presentation
document.addEventListener('DOMContentLoaded', function() {
    updateSlideCounter();
    updateNavButtons();
});

// Navigation functions
function nextSlide() {
    if (currentSlide < totalSlides) {
        document.getElementById(`slide-${currentSlide}`).classList.remove('active');
        currentSlide++;
        document.getElementById(`slide-${currentSlide}`).classList.add('active');
        updateSlideCounter();
        updateNavButtons();
    }
}

function prevSlide() {
    if (currentSlide > 1) {
        document.getElementById(`slide-${currentSlide}`).classList.remove('active');
        currentSlide--;
        document.getElementById(`slide-${currentSlide}`).classList.add('active');
        updateSlideCounter();
        updateNavButtons();
    }
}

// Go to specific slide
function goToSlide(slideNumber) {
    if (slideNumber >= 1 && slideNumber <= totalSlides) {
        document.getElementById(`slide-${currentSlide}`).classList.remove('active');
        currentSlide = slideNumber;
        document.getElementById(`slide-${currentSlide}`).classList.add('active');
        updateSlideCounter();
        updateNavButtons();
    }
}

// Update slide counter
function updateSlideCounter() {
    document.getElementById('current-slide').textContent = currentSlide;
    document.getElementById('total-slides').textContent = totalSlides;
}

// Update navigation button states
function updateNavButtons() {
    const prevBtn = document.getElementById('prev-btn');
    const nextBtn = document.getElementById('next-btn');
    
    prevBtn.disabled = (currentSlide === 1);
    nextBtn.disabled = (currentSlide === totalSlides);
}

// Keyboard navigation
document.addEventListener('keydown', function(event) {
    switch(event.key) {
        case 'ArrowRight':
        case ' ': // Space bar
            event.preventDefault();
            nextSlide();
            break;
        case 'ArrowLeft':
            event.preventDefault();
            prevSlide();
            break;
        case 'Home':
            event.preventDefault();
            goToSlide(1);
            break;
        case 'End':
            event.preventDefault();
            goToSlide(totalSlides);
            break;
    }
});

// Touch/swipe support for mobile
let touchStartX = 0;
let touchEndX = 0;

document.addEventListener('touchstart', function(event) {
    touchStartX = event.changedTouches[0].screenX;
});

document.addEventListener('touchend', function(event) {
    touchEndX = event.changedTouches[0].screenX;
    handleSwipe();
});

function handleSwipe() {
    const swipeThreshold = 50;
    const swipeDistance = touchEndX - touchStartX;
    
    if (Math.abs(swipeDistance) > swipeThreshold) {
        if (swipeDistance > 0) {
            // Swipe right - go to previous slide
            prevSlide();
        } else {
            // Swipe left - go to next slide
            nextSlide();
        }
    }
}

// Auto-save presentation state
function saveProgress() {
    localStorage.setItem('credora-presentation-slide', currentSlide);
}

// Load saved progress
function loadProgress() {
    const savedSlide = localStorage.getItem('credora-presentation-slide');
    if (savedSlide && savedSlide >= 1 && savedSlide <= totalSlides) {
        goToSlide(parseInt(savedSlide));
    }
}

// Save progress on slide change
setInterval(saveProgress, 1000);

// Load progress on page load
window.addEventListener('load', loadProgress);

// Presentation controls
function toggleFullscreen() {
    if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen();
    } else {
        document.exitFullscreen();
    }
}

// Add fullscreen toggle with F11
document.addEventListener('keydown', function(event) {
    if (event.key === 'F11') {
        event.preventDefault();
        toggleFullscreen();
    }
});